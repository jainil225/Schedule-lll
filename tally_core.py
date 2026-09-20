# -*- coding: utf-8 -*-
"""
tally_sch3.py  —  One-command Schedule III Financial Statement Generator
=========================================================================
Connects to TallyPrime → fetches all data → generates PDF instantly.
No intermediate JSON file needed (JSON also saved for reference).

USAGE:
  python tally_sch3.py --list-companies
  python tally_sch3.py --company "DEEP DEMO" --year 2026
  python tally_sch3.py --company "DEEP DEMO" --year 2026 --open
  python tally_sch3.py --company "DEEP DEMO" --year 2026 --port 9001

TALLY REQUIREMENTS (what client must enable):
  • F11 > Accounting Features  → Maintain Accounts       : Yes
  • F11 > Accounting Features  → Enable Bill-wise Entry  : Yes  (for Debtors/Creditors)
  • F11 > Inventory Features   → Maintain Inventory      : Yes  (for Stock)
  • F11 > Taxation             → Enable GST              : Yes
  • Gateway of Tally           → F12 > Configure         → Allow Browser Access for Reports : Yes
  • Tally must be running with HTTP server on port 9000 (default)
"""

import sys, io, re, json, argparse, os, subprocess, time
from datetime import datetime, date

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ── auto-install dependencies ─────────────────────────────────────────────────
def _pip(pkg):
    try: __import__(pkg)
    except ImportError:
        print(f"  Installing {pkg}...", end=" ", flush=True)
        subprocess.run([sys.executable, "-m", "pip", "install", pkg, "--quiet"], check=False)
        print("done")

_pip("requests"); _pip("reportlab")

import requests as req
import xml.etree.ElementTree as ET
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, PageBreak, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

TALLY_URL = "http://localhost:9000"   # overridden by --port arg

def set_tally_url(url):
    """Allow external callers (e.g. the GUI) to set the Tally endpoint."""
    global TALLY_URL
    TALLY_URL = url
    return TALLY_URL

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — TALLY XML HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def xe(t):
    return str(t).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def fix_enc(raw):
    raw  = re.sub(rb'<\?xml[^?]*\?>\s*', b'', raw)
    text = raw.decode("latin-1", errors="replace")
    for b, g in {'\x91':"'",'\x92':"'",'\x93':'"','\x94':'"','\x96':'-','\x97':'-'}.items():
        text = text.replace(b, g)
    def _san(m):
        try:
            cp = int(m.group(1)) if m.group(1) else int(m.group(2), 16)
            return m.group(0) if (cp in (9,10,13) or 0x20<=cp<=0xD7FF or 0xE000<=cp<=0xFFFD) else ""
        except: return ""
    text = re.sub(r'&#(\d+);|&#[xX]([0-9A-Fa-f]+);', _san, text)
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', text)
    return b'<?xml version="1.0" encoding="UTF-8"?>\n' + text.encode("utf-8", "replace")

def post(body, timeout=120):
    for attempt in range(3):
        try:
            r = req.post(TALLY_URL,
                         data=body.encode("utf-8"),
                         headers={"Content-Type": "text/xml; charset=utf-8"},
                         timeout=timeout)
            r.raise_for_status()
            return r.content
        except Exception as e:
            if attempt < 2:
                wait = 3 + attempt * 3
                print(f"\n    [retry {attempt+1}/2 in {wait}s — {e}]", end=" ", flush=True)
                time.sleep(wait)
            else:
                raise

def xml_to_dict(el, max_depth=6, depth=0):
    if depth > max_depth: return el.text
    d = {}
    if el.attrib: d["@attrs"] = dict(el.attrib)
    children = list(el)
    if not children: return (el.text or "").strip() or None
    seen = {}
    for ch in children:
        val = xml_to_dict(ch, max_depth, depth+1)
        tag = ch.tag
        if tag in seen:
            if not isinstance(seen[tag], list): seen[tag] = [seen[tag]]
            seen[tag].append(val)
        else:
            seen[tag] = val
    return seen

def safe_parse(raw):
    try:
        return xml_to_dict(ET.fromstring(fix_enc(raw)))
    except Exception as e:
        return {"__error__": str(e)}

def tally_date(d):
    if isinstance(d, str): d = datetime.strptime(d, "%Y-%m-%d").date()
    return d.strftime("%Y%m%d")

def fy_dates(year):
    """Given year=2026, return (2025-04-01, 2026-03-31) i.e. FY 2025-26."""
    end   = date(year, 3, 31)
    start = date(year-1, 4, 1)
    return str(start), str(end)

def find_list(obj, tag, results, depth=0):
    if depth > 6: return
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k.upper() == tag.upper() and isinstance(v, list):
                results.extend(v); return
            find_list(v, tag, results, depth+1)
    elif isinstance(obj, list):
        for item in obj: find_list(item, tag, results, depth+1)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — LIST COMPANIES
# ══════════════════════════════════════════════════════════════════════════════
def list_companies():
    body = """<ENVELOPE>
<HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Collection</TYPE><ID>MyCompanies</ID></HEADER>
<BODY><DESC>
<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>
<TDL><TDLMESSAGE>
<COLLECTION NAME="MyCompanies" ISMODIFY="No">
<TYPE>Company</TYPE><FETCH>Name</FETCH>
</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""
    try:
        raw  = post(body, timeout=15)
        root = ET.fromstring(fix_enc(raw))
        seen = set(); out = []
        skip = {"yes","no","true","false","tally","tallyprime","tallyerp",""}
        for el in root.iter():
            for src in [el.get("NAME"), el.text if el.tag == "NAME" else None]:
                t = (src or "").strip()
                if t and t.lower() not in seen and t.lower() not in skip and len(t) > 1:
                    seen.add(t.lower()); out.append(t)
        return out
    except Exception as e:
        print(f"  Error listing companies: {e}"); return []

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — ALL FETCH FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════
def fetch_trial_balance(company, fy_start, fy_end):
    """Primary: Tally's Export Data Trial Balance report.
    Fallback: Collection-based ledger fetch (works across FY periods)."""
    body = f"""<ENVELOPE>
<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
<BODY><EXPORTDATA><REQUESTDESC>
<REPORTNAME>Trial Balance</REPORTNAME>
<STATICVARIABLES>
  <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
  <SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>
  <SVFROMDATE>{tally_date(fy_start)}</SVFROMDATE>
  <SVTODATE>{tally_date(fy_end)}</SVTODATE>
  <EXPLODEFLAG>Yes</EXPLODEFLAG>
  <SVSHOWZEROBALANCE>No</SVSHOWZEROBALANCE>
</STATICVARIABLES>
</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>"""
    result = safe_parse(post(body, timeout=120))
    if isinstance(result, dict) and 'LINEERROR' in result:
        print(f"\n    Export Data failed, falling back to Collection TB…", end=" ", flush=True)
        return _fetch_tb_collection(company, fy_start, fy_end)
    return result

def _fetch_tb_collection(company, fy_start, fy_end):
    """Build TB from Collection API with date-range support.
    Uses CLOSINGBALANCE for current FY (correct) — OPENINGBALANCE was wrong.
    Falls back to OpeningBalance only if CLOSINGBALANCE is entirely zero."""
    body = f"""<ENVELOPE>
<HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Collection</TYPE><ID>TBLedgers</ID></HEADER>
<BODY><DESC>
<STATICVARIABLES>
  <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
  <SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>
  <SVFROMDATE>{tally_date(fy_start)}</SVFROMDATE>
  <SVTODATE>{tally_date(fy_end)}</SVTODATE>
</STATICVARIABLES>
<TDL><TDLMESSAGE>
<COLLECTION NAME="TBLedgers" ISMODIFY="No">
  <TYPE>Ledger</TYPE>
  <FETCH>Name,Parent,ClosingBalance,OpeningBalance</FETCH>
</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""
    ledger_raw = safe_parse(post(body, timeout=120))
    led_list = []
    find_list(ledger_raw, 'LEDGER', led_list)
    ledgers = []
    for item in led_list:
        if not isinstance(item, dict): continue
        nm = item.get('NAME', '')
        if isinstance(nm, dict):
            ln = nm if 'NAME.LIST' in nm else item.get('LANGUAGENAME.LIST', {})
            if isinstance(ln, dict):
                nl = ln.get('NAME.LIST', {})
                if isinstance(nl, dict):
                    nm = nl.get('NAME', '')
                    if isinstance(nm, list): nm = nm[0] if nm else ''
        if not nm or not isinstance(nm, str): continue
        parent = item.get('PARENT', '')
        # CLOSINGBALANCE is the period-end balance — correct for all date ranges
        try: cb = float(item.get('CLOSINGBALANCE', 0) or 0)
        except: cb = 0.0
        try: ob = float(item.get('OPENINGBALANCE', 0) or 0)
        except: ob = 0.0
        # Use CLOSINGBALANCE; fall back to OPENINGBALANCE if closing is zero
        # (some TallyPrime versions omit CLOSINGBALANCE for zero-movement ledgers)
        net = cb if cb != 0.0 else ob
        ledgers.append({'name': nm, 'parent': parent, 'closing': net})
    if not ledgers:
        return ledger_raw
    from collections import OrderedDict
    groups = OrderedDict()
    for led in sorted(ledgers, key=lambda x: x['parent']):
        par = led['parent']
        groups.setdefault(par, []).append(led)
    tb_names, tb_info = [], []
    for gname, children in groups.items():
        grp_dr = sum(l['closing'] for l in children if l['closing'] < 0)
        grp_cr = sum(l['closing'] for l in children if l['closing'] > 0)
        tb_names.append({'DSPDISPNAME': gname})
        tb_info.append({'DSPCLDRAMT': {'DSPCLDRAMTA': str(grp_dr)}, 'DSPCLCRAMT': {'DSPCLCRAMTA': str(grp_cr)}})
        for led in sorted(children, key=lambda x: x['name']):
            net = led['closing']
            tb_names.append({'DSPDISPNAME': led['name']})
            tb_info.append({'DSPCLDRAMT': {'DSPCLDRAMTA': str(net) if net < 0 else '0'},
                            'DSPCLCRAMT': {'DSPCLCRAMTA': str(net) if net > 0 else '0'}})
    print(f"OK ({len(tb_names)} rows)")
    return {'DSPACCNAME': tb_names, 'DSPACCINFO': tb_info}

def fetch_groups(company):
    body = f"""<ENVELOPE>
<HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Collection</TYPE><ID>AllGroups</ID></HEADER>
<BODY><DESC>
<STATICVARIABLES>
  <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
  <SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>
</STATICVARIABLES>
<TDL><TDLMESSAGE>
<COLLECTION NAME="AllGroups" ISMODIFY="No">
  <TYPE>Group</TYPE>
  <FETCH>Name,Parent,NatureOfGroup,IsRevenue,AffectsGrossProfit,IsAddable,MasterId</FETCH>
</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""
    return safe_parse(post(body, timeout=60))

def fetch_stock_items(company, fy_end):
    body = f"""<ENVELOPE>
<HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Collection</TYPE><ID>StockColl</ID></HEADER>
<BODY><DESC>
<STATICVARIABLES>
  <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
  <SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>
  <SVTODATE>{tally_date(fy_end)}</SVTODATE>
</STATICVARIABLES>
<TDL><TDLMESSAGE>
<COLLECTION NAME="StockColl" ISMODIFY="No">
  <TYPE>StockItem</TYPE>
  <FETCH>Name,Parent,BaseUnits,StandardCost,StandardPrice,
         OpeningBalance,OpeningRate,OpeningValue,
         ClosingBalance,ClosingRate,ClosingValue,
         GSTDetails,HSNDetails,TaxClassificationName,MasterId,Category</FETCH>
</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""
    return safe_parse(post(body, timeout=120))

def fetch_outstanding_bills(company, fy_end):
    body = f"""<ENVELOPE>
<HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Collection</TYPE><ID>OutstandingColl</ID></HEADER>
<BODY><DESC>
<STATICVARIABLES>
  <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
  <SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>
  <SVTODATE>{tally_date(fy_end)}</SVTODATE>
</STATICVARIABLES>
<TDL><TDLMESSAGE>
<COLLECTION NAME="OutstandingColl" ISMODIFY="No">
  <TYPE>Ledger</TYPE>
  <FETCH>Name,Parent,OpeningBalance,ClosingBalance,
         IsBillwiseOn,BillCreditPeriod,CreditLimit,
         BillAllocations,PartyGSTIN</FETCH>
</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""
    return safe_parse(post(body, timeout=120))

def fetch_company_info(company):
    body = f"""<ENVELOPE>
<HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Collection</TYPE><ID>CompanyInfo</ID></HEADER>
<BODY><DESC>
<STATICVARIABLES>
  <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
  <SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>
</STATICVARIABLES>
<TDL><TDLMESSAGE>
<COLLECTION NAME="CompanyInfo" ISMODIFY="No">
  <TYPE>Company</TYPE>
  <FETCH>Name,FormalName,Address,StateName,CountryName,PinCode,
         PhoneNumber,Email,Website,GSTRegistrationDate,
         GSTNumber,PANItNumber,CINNumber,BookFrom,IsMSME,MSMEType</FETCH>
</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""
    return safe_parse(post(body, timeout=30))

def check_payroll(company):
    body = f"""<ENVELOPE>
<HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Collection</TYPE><ID>PayrollCheck</ID></HEADER>
<BODY><DESC>
<STATICVARIABLES>
  <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
  <SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>
</STATICVARIABLES>
<TDL><TDLMESSAGE>
<COLLECTION NAME="PayrollCheck" ISMODIFY="No">
  <TYPE>Company</TYPE>
  <FETCH>Name,IsPayrollEnabled,UsePayroll</FETCH>
</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""
    try:
        text = post(body, timeout=15).decode("latin-1","replace").upper()
        return "ISPAYROLLENABLED>YES" in text or "USEPAYROLL>YES" in text
    except: return False

def fetch_native_pnl(company, fy_start, fy_end):
    """Fast closing stock fetch — never hangs Tally.
    All strategies have 8s timeout so large companies fail fast and fall through.
    Returns dict with closing_stock, opening_stock, report_name.
    """
    _empty = {"closing_stock": None, "opening_stock": None, "report_name": None}
    _from  = tally_date(fy_start)
    _to    = tally_date(fy_end)

    # Strategy 1: TDL — Stock-in-Hand ledger closing balances (TYPE A companies)
    try:
        b = (f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
             f"<TYPE>Collection</TYPE><ID>NP1</ID></HEADER><BODY><DESC>"
             f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
             f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
             f"<SVFROMDATE>{_from}</SVFROMDATE><SVTODATE>{_to}</SVTODATE>"
             f"</STATICVARIABLES><TDL><TDLMESSAGE>"
             f"<COLLECTION NAME=\"NP1\" ISMODIFY=\"No\"><TYPE>Ledger</TYPE>"
             f"<BELONGSTO>Stock-in-Hand</BELONGSTO>"
             f"<FETCH>Name,ClosingBalance,OpeningBalance</FETCH>"
             f"</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>")
        items = []
        find_list(safe_parse(post(b, timeout=8)), 'LEDGER', items)
        cs = sum(abs(fv(l.get('CLOSINGBALANCE',0))) for l in items if isinstance(l,dict))
        os_ = sum(abs(fv(l.get('OPENINGBALANCE',0))) for l in items if isinstance(l,dict))
        if cs > 0:
            print(f"  Native PnL (S1 TDL-ledger): CS={fmt_indian(cs)}")
            return {"closing_stock": cs, "opening_stock": os_ or None, "report_name": "S1"}
    except: pass

    # Strategy 2: Stock Items ClosingValue with SVFROMDATE (TYPE B companies — Aeliya)
    try:
        b = (f"<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
             f"<TYPE>Collection</TYPE><ID>NP2</ID></HEADER><BODY><DESC>"
             f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
             f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
             f"<SVFROMDATE>{_from}</SVFROMDATE><SVTODATE>{_to}</SVTODATE>"
             f"</STATICVARIABLES><TDL><TDLMESSAGE>"
             f"<COLLECTION NAME=\"NP2\" ISMODIFY=\"No\"><TYPE>StockItem</TYPE>"
             f"<FETCH>Name,ClosingBalance,ClosingRate,ClosingValue</FETCH>"
             f"</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>")
        items = []
        find_list(safe_parse(post(b, timeout=8)), 'STOCKITEM', items)
        cs = sum(abs(fv(s.get('CLOSINGVALUE',0))) for s in items if isinstance(s,dict))
        if cs > 0:
            print(f"  Native PnL (S2 StockItems+date): CS={fmt_indian(cs)}")
            return {"closing_stock": cs, "opening_stock": None, "report_name": "S2"}
    except: pass

    # Strategy 3: P&L report — 8s timeout so it fails fast for large companies
    def _amt(text):
        try: return float((text or "").strip().replace(",",""))
        except: return 0.0

    for rname in ["Profit and Loss", "Trading and Profit and Loss",
                  "Profit & Loss", "Profit & Loss A/c"]:
        try:
            b = (f"<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>"
                 f"<BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>{xe(rname)}</REPORTNAME>"
                 f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
                 f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
                 f"<SVFROMDATE>{_from}</SVFROMDATE><SVTODATE>{_to}</SVTODATE>"
                 f"<EXPLODEFLAG>Yes</EXPLODEFLAG></STATICVARIABLES>"
                 f"</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>")
            raw = post(b, timeout=8)
            root = ET.fromstring(fix_enc(raw))
            if sum(1 for _ in root.iter("DSPACCNAME")) == 0: continue
            result = {"closing_stock": None, "opening_stock": None, "report_name": rname}
            all_e  = list(root.iter())
            for i, el in enumerate(all_e):
                if el.tag != "DSPACCNAME": continue
                ne = el.find("DSPDISPNAME")
                nm = (ne.text or "").strip() if ne is not None else (el.text or "").strip()
                nl = nm.lower()
                isc = "closing stock" in nl or "closing inventor" in nl or ("less" in nl and "stock" in nl)
                iso = "opening stock" in nl or "opening inventor" in nl
                if not isc and not iso: continue
                amt = 0.0
                par = next((p for p in all_e if el in list(p)), None)
                if par:
                    for info in par.iter("DSPACCINFO"):
                        for tg in ("DSPCLDRAMT","DSPCLCRAMT"):
                            s2 = info.find(tg)
                            if s2:
                                for at in ("DSPCLDRAMTA","DSPCLCRAMTA"):
                                    a = s2.find(at)
                                    if a is not None and a.text: amt += _amt(a.text)
                        break
                    if amt == 0:
                        pc = list(par)
                        for ci,ch in enumerate(pc):
                            if ch is el:
                                for sib in pc[ci+1:ci+4]:
                                    if sib.tag == "PLAMT":
                                        s3 = sib.find("PLSUBAMT")
                                        if s3 is not None and s3.text: amt += _amt(s3.text)
                                        break
                                break
                if amt == 0:
                    for fw in all_e[i+1:i+6]:
                        if fw.tag == "PLAMT":
                            s4 = fw.find("PLSUBAMT")
                            if s4 is not None and s4.text: amt += _amt(s4.text)
                            break
                        if fw.tag == "DSPACCNAME": break
                v = abs(amt) if amt else None
                if v:
                    if isc: result["closing_stock"] = v
                    elif iso: result["opening_stock"] = v
            if result["closing_stock"] or result["opening_stock"]:
                return result
        except: continue
    return _empty

def fetch_all(company, year, save_json=True,
              from_date=None, to_date=None):
    """
    Fetch all Tally data for `company`.
    - If from_date / to_date are supplied (YYYY-MM-DD strings), they are used
      directly as the period — allowing sub-FY or multi-FY fetches.
    - Otherwise the standard Indian FY for `year` is used (Apr–Mar).
    """
    if from_date and to_date:
        fy_start, fy_end = str(from_date), str(to_date)
    else:
        fy_start, fy_end = fy_dates(year)

    print(f"\n{'═'*65}")
    print(f"  Company  : {company}")
    print(f"  FY       : {fy_start}  →  {fy_end}  (Year {year-1}-{str(year)[2:]})")
    print(f"  Tally    : {TALLY_URL}")
    print(f"{'═'*65}")

    # Test connection
    try:
        req.post(TALLY_URL, data=b"<ENVELOPE/>", timeout=4)
        print("  Tally connection : OK")
    except Exception as e:
        print(f"\n  ERROR: Tally not reachable at {TALLY_URL}")
        print(f"  → {e}")
        print(f"\n  Make sure:")
        print(f"    1. TallyPrime is open and running")
        print(f"    2. The company '{company}' is loaded")
        print(f"    3. F12 > Configure > Enable Browser Access for Reports = Yes")
        print(f"    4. Port matches (default 9000, use --port to change)")
        sys.exit(1)

    data = {}
    # Previous FY dates (for previous year column in BS / PnL)
    from datetime import date as _date
    _py_end   = str(_date(year-1, 3, 31))   # previous year end
    _py_start = str(_date(year-2, 4, 1))    # previous year start

    steps = [
        ("Trial Balance",         lambda: fetch_trial_balance(company, fy_start, fy_end)),
        ("Trial Balance PY",      lambda: fetch_trial_balance(company, _py_start, _py_end)),
        ("Group Hierarchy",       lambda: fetch_groups(company)),
        ("Stock Items",           lambda: fetch_stock_items(company, fy_end)),
        ("Outstanding Bills",     lambda: fetch_outstanding_bills(company, fy_end)),
        ("Outstanding Bills PY",  lambda: fetch_outstanding_bills(company, _py_end)),
        ("Company Info",          lambda: fetch_company_info(company)),
        # Native PnL disabled — full P&L report hangs large companies (Aeliya etc.)
        # Closing stock is handled directly in parse_data via Stock Items ClosingValue.
    ]

    for label, fn in steps:
        print(f"  Fetching {label:25s}...", end=" ", flush=True)
        t0 = time.time()
        try:
            time.sleep(0.5)   # small gap — Tally drops connection if hammered
            data[label] = fn()
            print(f"OK  ({time.time()-t0:.1f}s)")
        except Exception as e:
            print(f"FAILED\n    → {e}")
            data[label] = {"__error__": str(e)}

    # Save JSON for reference
    if save_json:
        safe = re.sub(r'[^\w]', '_', company[:30])
        jf   = f"tally_raw_{safe}_{fy_end}.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump({
                "company": company, "fy_end": fy_end, "fy_start": fy_start,
                "generated_at": datetime.now().isoformat(),
                "tally_url": TALLY_URL, "data": data
            }, f, indent=2, default=str)
        print(f"\n  Raw JSON saved → {jf}")

    return data, fy_start, fy_end

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4B — DYNAMIC COMPANY TYPE DETECTOR
# ══════════════════════════════════════════════════════════════════════════════
def detect_company_type(TB: dict) -> dict:
    """
    Dynamically detect company type from Trial Balance group balances.

    TYPE DETECTION LOGIC:
      Stock-in-Hand / Opening Stock group has non-zero balance → has_stock = True
        + Purchase Accounts + Direct Expenses → manufacturing
        + Purchase Accounts only              → trading (with inventory)
      Purchase Accounts, NO stock             → trading (no inventory, e.g. Vianci)
      Sales only, NO purchases                → service company
      Everything else                         → mixed

    EFFECT ON P&L:
      has_stock = False → opening_stock = 0, closing_stock = 0
                        → gross_profit = Sales + Direct Inc + Direct Exp (signed)
                        → Trading Account section hidden in UI
    EFFECT ON BS:
      has_stock = False → no Inventory / Closing Stock line in Current Assets
                        → BS Difference = 0 (no phantom stock entry)

    Returns dict:
      { has_stock, has_inventory, type, closing_stock }
    """
    stock_groups = {
        'Stock-in-Hand', 'Opening Stock', 'Closing Stock',
        'Stock in Trade', 'Inventory',
    }
    stock_val = sum(
        abs(v) for k, v in TB.items()
        if any(sg.lower() in k.lower() for sg in stock_groups)
    )
    has_purchases = abs(TB.get('Purchase Accounts', 0)) > 1
    has_dir_exp   = abs(TB.get('Direct Expenses', 0)) > 1
    has_dir_inc   = abs(TB.get('Direct Incomes', 0)) > 1
    has_stock     = stock_val > 100   # ₹100 threshold ignores rounding noise

    has_sales     = abs(TB.get('Sales Accounts', 0)) > 1
    has_ind_inc   = abs(TB.get('Indirect Incomes', 0)) > 1
    has_ind_exp   = abs(TB.get('Indirect Expenses', 0)) > 1
    has_any_activity = has_sales or has_purchases or has_dir_exp or has_dir_inc or has_ind_inc or has_ind_exp

    # FIX: Detect dormant/shell companies — zero income AND zero expenses for the period.
    # These must not be forced into "service" type which would try to render P&L figures.
    # Dormant companies have only BS items: capital, loans, creditors, bank, P&L b/f opening loss.
    if not has_any_activity and not has_stock:
        ctype = "dormant"
    elif has_stock and has_purchases:
        ctype = "manufacturing" if has_dir_exp else "trading"
    elif has_purchases and not has_stock:
        ctype = "trading"       # trading without stock ledgers (e.g. Vianci)
    elif not has_purchases and has_dir_inc:
        # Job work / commission agent — Direct Incomes but no purchase accounts.
        # Treat as service for stock purposes but retain direct income lines on P&L.
        ctype = "service"
    elif not has_purchases:
        ctype = "service"
    else:
        ctype = "mixed"

    return {
        "has_stock"    : has_stock,
        "has_inventory": has_stock,
        "type"         : ctype,
        "closing_stock": stock_val if has_stock else 0.0,
        "is_dormant"   : ctype == "dormant",
    }


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — DATA PARSER  (raw API data → clean Python dict)
# ══════════════════════════════════════════════════════════════════════════════
def parse_data(company, data, fy_start, fy_end):

    # ── helpers ───────────────────────────────────────────────────────────────
    def gname(g):
        ln = g.get('LANGUAGENAME.LIST', {})
        if isinstance(ln, dict):
            nl = ln.get('NAME.LIST', {})
            if isinstance(nl, dict):
                nm = nl.get('NAME', '')
                return (nm[0] if isinstance(nm, list) else nm) or ''
        return ''

    def fv(s):
        if s is None: return 0.0
        if isinstance(s, dict):
            for key in ('AMOUNT', 'DSPCLDRAMTA', 'DSPCLCRAMTA', 'RatePerUnit'):
                if key in s: return fv(s[key])
            for v in s.values():
                try: return float(v)
                except: pass
            return 0.0
        try: return float(str(s).replace(',', ''))
        except: return 0.0

    # ── Trial Balance ─────────────────────────────────────────────────────────
    tb_raw = data.get("Trial Balance", {})
    TB = {}
    TB_ORDER = []   # preserve Tally's original order (group followed by its children)

    # DSPACCNAME/DSPACCINFO may be at root or nested inside intermediate tags
    def _find_key(obj, key, depth=0):
        if depth > 8 or not isinstance(obj, dict): return None
        if key in obj: return obj[key]
        for v in obj.values():
            r = _find_key(v, key, depth+1)
            if r is not None: return r
        return None

    tb_names = _find_key(tb_raw, 'DSPACCNAME') or []
    tb_info  = _find_key(tb_raw, 'DSPACCINFO') or []
    if not isinstance(tb_names, list): tb_names = [tb_names]
    if not isinstance(tb_info, list):  tb_info  = [tb_info]
    for n, info in zip(tb_names, tb_info):
        name = n.get('DSPDISPNAME', '') if isinstance(n, dict) else ''
        if not name: continue
        dr = fv(info.get('DSPCLDRAMT',{}).get('DSPCLDRAMTA') if isinstance(info,dict) else 0)
        cr = fv(info.get('DSPCLCRAMT',{}).get('DSPCLCRAMTA') if isinstance(info,dict) else 0)
        TB[name] = dr + cr   # negative = DR (asset/expense), positive = CR (liability/income)
        TB_ORDER.append(name)

    if not TB:
        print("\n  ERROR: Trial Balance returned 0 rows.")
        sys.exit(1)

    # ── Groups ────────────────────────────────────────────────────────────────
    grp_raw = data.get("Group Hierarchy", {})
    groups_list = []
    find_list(grp_raw, 'GROUP', groups_list)

    PARENT     = {}
    IS_REVENUE = {}
    AFFECTS_GP = {}
    for g in groups_list:
        if not isinstance(g, dict): continue
        nm = gname(g)
        if nm:
            PARENT[nm]     = g.get('PARENT', '')
            IS_REVENUE[nm] = g.get('ISREVENUE', 'No') == 'Yes'
            AFFECTS_GP[nm] = g.get('AFFECTSGROSSPROFIT', 'No') == 'Yes'

    def get_primary(name, depth=0):
        if depth > 12: return name
        p = PARENT.get(name)
        if not p or p == 'Primary': return name
        return get_primary(p, depth+1)

    # ── Stock — from Stock-in-Hand LEDGERS (matches Tally P&L exactly) ─────────
    # These ledgers under "Stock-in-Hand" group give CORRECT closing stock.
    # e.g. Stock of R.M. (Rajpur), Stock of F.G. (Rajpur), Stock WIP etc.
    # This matches what Tally P&L shows as Closing Stock (e.g. 91,28,20,229)
    # The Stock Items API returns per-item physical values which differ.
    ob_stk_raw = data.get("Outstanding Bills", {})
    ob_stk_list = []
    find_list(ob_stk_raw, 'LEDGER', ob_stk_list)

    stk_ledgers   = []   # [{name, closing, opening}] — for P&L breakdown
    stock_by_cat  = {}   # category -> [{name, closing_val, opening_val}] for Notes
    for item in ob_stk_list:
        if not isinstance(item, dict): continue
        if item.get('PARENT', '') == 'Stock-in-Hand':
            nm = gname(item) or ''
            if isinstance(nm, list): nm = nm[0] if nm else ''
            cb = abs(fv(item.get('CLOSINGBALANCE', 0)))
            ob = abs(fv(item.get('OPENINGBALANCE', 0)))
            stk_ledgers.append({'name': nm, 'closing': cb, 'opening': ob})
            stock_by_cat.setdefault('Stock-in-Hand', []).append({
                'name': nm, 'closing_val': -cb, 'opening_val': -ob,
                'unit': '', 'closing_qty': '',
            })

    closing_stock_abs = sum(i['closing'] for i in stk_ledgers)

    # ── Previous year stock data ───────────────────────────────────────────────
    ob_py_raw  = data.get("Outstanding Bills PY", {})
    ob_py_list = []
    find_list(ob_py_raw, 'LEDGER', ob_py_list)
    stk_ledgers_py = []
    stk_raw_py     = []
    for item in ob_py_list:
        if not isinstance(item, dict): continue
        if item.get('PARENT', '') == 'Stock-in-Hand':
            nm = gname(item) or ''
            if isinstance(nm, list): nm = nm[0] if nm else ''
            cb = abs(fv(item.get('CLOSINGBALANCE', 0)))
            ob = abs(fv(item.get('OPENINGBALANCE', 0)))
            stk_ledgers_py.append({'name': nm, 'closing': cb, 'opening': ob})
            stk_raw_py.append(item)
    closing_stock_abs_py = sum(i['closing'] for i in stk_ledgers_py)

    # ── Stock figures — universal auto-detecting logic ────────────────────────
    #
    # Tally has TWO types of companies w.r.t. Stock-in-Hand LEDGERS:
    #
    #  TYPE A — "Ledger-based stock" (e.g. DEEP DEMO / manufacturing):
    #    Stock-in-Hand group contains LEDGERS (e.g. "Stock of R.M. (Rajpur)").
    #    These ledgers carry real CLOSINGBALANCE and OPENINGBALANCE values.
    #    OB fetch (as at fy_end):
    #      CLOSINGBALANCE = stock as at fy_end  → Closing Stock
    #      OPENINGBALANCE = stock as at fy_start → Opening Stock
    #
    #  TYPE B — "Item-based stock" (e.g. Trulean / trading):
    #    Stock-in-Hand group may have a single ledger OR none, and the
    #    CLOSINGBALANCE on that ledger is 0 (Tally stores stock physically
    #    via Stock Items, not as a ledger balance).
    #    In this case:
    #      TB 'Opening Stock' row          → Opening Stock (DR entry at year start)
    #      Stock Items API ClosingValue     → Closing Stock (physical inventory)
    #      OB_PY fetch CLOSINGBALANCE      → Closing Stock fallback (prev year actual closing)
    #
    # DETECTION: if sum(CLOSINGBALANCE) from current OB fetch > 0 → TYPE A
    #            else → TYPE B

    # ── Opening Stock (P&L Dr side) ──────────────────────────────────────────
    # TB 'Opening Stock' row is the most reliable source for BOTH types.
    # Tally always posts an 'Opening Stock' DR entry at year start from Stock Items.
    opening_stock_abs = abs(TB.get('Opening Stock', 0))
    if opening_stock_abs == 0:
        # Fallback: sum of OPENINGBALANCE from current OB fetch (works for Type A)
        opening_stock_abs = sum(i['opening'] for i in stk_ledgers)

    # ── Detect company type: zero stock for service AND trading-without-stock ──
    # TWO scenarios that must produce zero inventory:
    #
    #  A) Services company  — no purchases, no stock ledgers, no Opening Stock.
    #     e.g. consulting firm, NGO, CA firm.
    #
    #  B) Trading-without-stock-ledgers — HAS purchases, BUT Tally is configured
    #     WITHOUT "Maintain Inventory" OR items are booked directly to expense
    #     ledgers (no Stock-in-Hand group ledgers, no Opening Stock in TB).
    #     e.g. Vianci Enterprise — purchases Rs.50 L but zero stock ledgers.
    #     Without this guard the TYPE B else-branch fires, falls through to the
    #     _py_closing fallback and assigns stale PY stock → phantom inventory
    #     on BS and inflated gross_profit in P&L.
    #
    # DETECTION: no stock ledger balances AND no Opening Stock row in TB.
    _co_type = detect_company_type(TB)
    _no_stk_ledgers          = not stk_ledgers
    _no_opening_stock        = opening_stock_abs == 0
    _is_services             = (_co_type["type"] == "service") and _no_stk_ledgers and _no_opening_stock
    _is_trading_no_inventory = (_co_type["type"] == "trading") and _no_stk_ledgers and _no_opening_stock

    if _is_services or _is_trading_no_inventory:
        # Zero stock everywhere — no Inventory line on BS, no stock terms in P&L
        bs_closing_stock  = 0.0
        stk_ledgers       = []
        opening_stock_abs = 0.0
        _label = "TRADING-NO-INVENTORY" if _is_trading_no_inventory else "SERVICE"
        print(f"  Stock: {_label} (no stock ledgers, no Opening Stock) -> bs_closing_stock=0")

    elif closing_stock_abs > 0:
        # TYPE A: ledgers carry real balances → use CLOSINGBALANCE directly.
        # BUT: Tally P&L may show a different closing stock figure (from Native PnL)
        # especially for custom date ranges where CLOSINGBALANCE = as-at-today
        # but Native PnL closing stock = as-at-period-end. Prefer Native PnL when available.
        native_pnl_a = data.get("Native PnL", {})
        _native_closing_a = native_pnl_a.get("closing_stock") if isinstance(native_pnl_a, dict) else None
        if _native_closing_a and _native_closing_a > 0 and abs(_native_closing_a - closing_stock_abs) > 1.0:
            # Native P&L value differs from ledger sum → use Native P&L (matches Tally P&L display)
            bs_closing_stock = _native_closing_a
            print(f"  Stock TYPE A (native P&L override): closing={fmt_indian(bs_closing_stock)} "
                  f"vs ledger={fmt_indian(closing_stock_abs)}")
        else:
            bs_closing_stock = closing_stock_abs

    else:
        # TYPE B: ledger closing is zero → use Tally's native P&L first,
        # then Stock Items API as fallback.
        native_pnl = data.get("Native PnL", {})
        _native_closing = native_pnl.get("closing_stock") if isinstance(native_pnl, dict) else None

        stk_raw = data.get("Stock Items", {})
        stk_items_list = []
        find_list(stk_raw, 'STOCKITEM', stk_items_list)
        _items_closing = sum(
            abs(fv(si.get('CLOSINGVALUE', 0)))
            for si in stk_items_list if isinstance(si, dict)
        )

        if _native_closing and _native_closing > 0:
            bs_closing_stock = _native_closing
            print(f"  Stock TYPE B (native P&L): closing={fmt_indian(bs_closing_stock)}")
        elif _items_closing > 0:
            bs_closing_stock = _items_closing
        else:
            _py_closing = sum(abs(fv(i.get('CLOSINGBALANCE', 0)))
                              for i in stk_raw_py if isinstance(i, dict))
            bs_closing_stock = _py_closing if _py_closing > 0 else opening_stock_abs

        if _items_closing > 0:
            _items_by_parent = {}
            for si in stk_items_list:
                if not isinstance(si, dict): continue
                cv = abs(fv(si.get('CLOSINGVALUE', 0)))
                if cv == 0: continue
                par = si.get('PARENT', 'Stock-in-Hand')
                _items_by_parent.setdefault(par, 0)
                _items_by_parent[par] += cv
            for sl in stk_ledgers:
                grp_val = _items_by_parent.get(sl['name'], 0)
                sl['closing'] = grp_val if grp_val > 0 else bs_closing_stock
            if not stk_ledgers:
                stk_ledgers.append({'name': 'Stock in Hand', 'closing': bs_closing_stock, 'opening': opening_stock_abs})
        elif not stk_ledgers:
            stk_ledgers.append({'name': 'Stock in Hand', 'closing': bs_closing_stock, 'opening': opening_stock_abs})

        _stk_sum = sum(sl['closing'] for sl in stk_ledgers)
        if _stk_sum > 0 and abs(_stk_sum - bs_closing_stock) > 1.0:
            _ratio = bs_closing_stock / _stk_sum
            for sl in stk_ledgers:
                sl['closing'] = round(sl['closing'] * _ratio, 2)

    # Ensure stk_ledgers breakdown sums correctly to bs_closing_stock
    _check_sum = sum(sl['closing'] for sl in stk_ledgers)
    if abs(_check_sum - bs_closing_stock) > 1.0 and len(stk_ledgers) == 1:
        stk_ledgers[0]['closing'] = bs_closing_stock

    # Rebuild stock_by_cat with final resolved closing values (used in Notes)
    stock_by_cat = {}
    for sl in stk_ledgers:
        if sl['closing'] > 0 or sl['opening'] > 0:
            stock_by_cat.setdefault('Stock-in-Hand', []).append({
                'name': sl['name'],
                'closing_val': -sl['closing'],
                'opening_val': -sl['opening'],
                'unit': '', 'closing_qty': '',
            })

    # ── Outstanding Bills ─────────────────────────────────────────────────────
    ob_raw = data.get("Outstanding Bills", {})
    ob_list = []
    find_list(ob_raw, 'LEDGER', ob_list)
    debtors_list   = []
    creditors_list = []
    for item in ob_list:
        if not isinstance(item, dict): continue
        cb     = fv(item.get('CLOSINGBALANCE', 0))
        if cb == 0: continue
        nm     = gname(item) or item.get('NAME', '')
        parent = item.get('PARENT', '')
        prim   = get_primary(parent)
        # Walk up the parent chain to catch nested sub-groups like
        # "Trade Receivables > Sundry Debtors > PartyName"
        _is_debtor   = ('Debtor' in parent or 'debtor' in parent.lower() or
                        prim in ('Sundry Debtors', 'SUNDRY DEBTORS') or
                        'debtor' in prim.lower() or 'receivable' in prim.lower())
        _is_creditor = (('Creditor' in parent or 'creditor' in parent.lower()) and
                        (prim in ('Current Liabilities', 'Sundry Creditors') or
                         'creditor' in prim.lower() or 'payable' in prim.lower()))
        if _is_debtor:
            debtors_list.append({'name': nm, 'parent': parent, 'balance': cb})
        elif _is_creditor:
            creditors_list.append({'name': nm, 'parent': parent, 'balance': cb})

    # ── Company Info ──────────────────────────────────────────────────────────
    co_raw  = data.get("Company Info", {})
    co_list = []
    find_list(co_raw, 'COMPANY', co_list)
    co_info = {}
    for c in co_list:
        if isinstance(c, dict) and c.get('NAME','').upper() == company.upper():
            co_info = c; break
    if not co_info and co_list:
        co_info = co_list[0] if isinstance(co_list[0], dict) else {}

    # ══════════════════════════════════════════════════════════════════════════
    # DYNAMIC GROUP CLASSIFIER
    # Tally TB is ordered: parent group row, then all its children immediately after.
    # We use positional extraction to find children of any group.
    # ALL primary Tally groups must be in TOP_GROUPS so extraction stops correctly.
    # ══════════════════════════════════════════════════════════════════════════

    # ALL known Tally primary groups — positional extraction stops at any of these
    TOP_GROUPS = {
        'Capital Account', 'Loans (Liability)', 'Current Liabilities',
        'Fixed Assets', 'Current Assets', 'Investments',
        'Misc. Expenses (ASSET)', 'Branch / Divisions', 'Suspense A/c',
        'Sales Accounts', 'Purchase Accounts', 'Direct Expenses',
        'Direct Incomes', 'Indirect Incomes', 'Indirect Expenses',
        'Profit & Loss A/c',
    }

    def tb_children_of(group_name):
        """
        Extract TB rows that are direct children of group_name using positional order.
        Stops as soon as any other primary group is encountered.
        Returns list of (name, tb_value) — includes zero-value rows for completeness;
        callers should filter zeros as needed.
        """
        if group_name not in TB_ORDER:
            return []
        idx = TB_ORDER.index(group_name)
        result = []
        for j in range(idx + 1, len(TB_ORDER)):
            candidate = TB_ORDER[j]
            # Stop at any other primary/top-level group
            if candidate in TOP_GROUPS or PARENT.get(candidate, '') == 'Primary':
                break
            val = TB.get(candidate, 0)
            result.append((candidate, val))
        return result

    # Build children map from groups hierarchy (for sub-groups known to Tally)
    children = {}
    for name in TB:
        par = PARENT.get(name, '')
        children.setdefault(par, []).append(name)

    def direct_children(group):
        """Direct children from Group Hierarchy (sub-groups only)."""
        kids = children.get(group, [])
        return sorted(kids, key=lambda n: abs(TB.get(n, 0)), reverse=True)

    def child_lines(group):
        return {n: TB[n] for n in direct_children(group) if TB.get(n, 0) != 0}

    # ── P&L Top-level groups ─────────────────────────────────────────────────
    sales_grp   = TB.get('Sales Accounts', 0)
    purch_grp   = TB.get('Purchase Accounts', 0)
    open_s_grp  = TB.get('Opening Stock', 0)
    dir_inc_grp = TB.get('Direct Incomes', 0)
    dir_exp_grp = TB.get('Direct Expenses', 0)
    ind_inc_grp = TB.get('Indirect Incomes', 0)
    ind_exp_grp = TB.get('Indirect Expenses', 0)
    pnl_bf      = TB.get('Profit & Loss A/c', 0)

    # ── Company-type-aware P&L formula ───────────────────────────────────────
    # TB sign: CR=positive (income/liability), DR=negative (expense/asset).
    # Direct addition is sign-correct for ALL cases:
    #   expense DR (negative) subtracts → correct
    #   expense CR (refund > charge)    → adds → correct
    #
    # STOCK TERMS must only be included for companies that actually carry stock.
    # For service companies and trading-without-stock-ledger companies (e.g. Vianci)
    # bs_closing_stock = 0 and opening_stock_abs = 0 (enforced above).
    # Including them anyway is harmless numerically when both are 0, BUT the
    # formula is documented here explicitly per type to avoid future confusion.
    #
    #   Manufacturing / Trading WITH stock:
    #     GP = Sales + Dir.Inc + Closing Stock + Purchases + (-Opening Stock) + Dir.Exp
    #
    #   Trading WITHOUT stock ledgers (e.g. Vianci):
    #     GP = Sales + Dir.Inc + Purchases + Dir.Exp
    #     (bs_closing_stock=0 and opening_stock_abs=0 already → same formula, no stock terms)
    #
    #   Service / NGO:
    #     GP = Sales + Dir.Inc + Dir.Exp   (no purchases, no stock)
    #
    # All three resolve to the same expression because the stock terms are zero
    # for non-stock companies. The guard above ensures bs_closing_stock=0 and
    # opening_stock_abs=0 before we reach here.
    gross_profit = (sales_grp + dir_inc_grp + bs_closing_stock
                    + purch_grp + (-opening_stock_abs) + dir_exp_grp)
    net_profit   = gross_profit + ind_inc_grp + ind_exp_grp

    # Diagnostic: log computed P&L for verification against Tally display
    print(f"  P&L: sales={fmt_indian(abs(sales_grp))}  purch={fmt_indian(abs(purch_grp))}"
          f"  dir_exp={fmt_indian(abs(dir_exp_grp))}  ind_exp={fmt_indian(abs(ind_exp_grp))}"
          f"  cs={fmt_indian(bs_closing_stock)}  os={fmt_indian(opening_stock_abs)}"
          f"  GP={fmt_indian(gross_profit)}  NP={fmt_indian(net_profit)}"
          f"  type={_co_type['type']}")

    # Dynamic child lines — use positional extraction for all groups
    sales_lines    = tb_children_of('Sales Accounts')
    purch_lines    = tb_children_of('Purchase Accounts')
    dir_inc_lines  = tb_children_of('Direct Incomes')
    dir_exp_lines  = tb_children_of('Direct Expenses')
    ind_inc_lines  = tb_children_of('Indirect Incomes')
    ind_exp_lines  = tb_children_of('Indirect Expenses')

    # ── BALANCE SHEET — positional extraction ────────────────────────────────
    cap_lines    = tb_children_of('Capital Account')
    loans_lines  = tb_children_of('Loans (Liability)')
    cl_lines     = tb_children_of('Current Liabilities')
    fa_lines_all = tb_children_of('Fixed Assets')

    # CA lines: Replace TB 'Opening Stock' row with Closing Stock at period-end.
    # bs_closing_stock is already zeroed for service/trading-without-inventory.
    # When it is zero, the 'Opening Stock' row is simply dropped (not replaced)
    # so no phantom stock line appears in CA for Vianci-type companies.
    _ca_raw  = tb_children_of('Current Assets')
    _has_os  = any(_nm == 'Opening Stock' for _nm, _val in _ca_raw)
    ca_lines = []
    for _nm, _val in _ca_raw:
        if _nm == 'Opening Stock' and _has_os:
            if bs_closing_stock > 0:
                # Stock company: replace Opening Stock row with Closing Stock value
                ca_lines.append(('Closing Stock', -bs_closing_stock))
                # Per-ledger breakdown only when multiple ledgers (avoids redundant sub-line)
                _nonzero_stk = [_sl for _sl in stk_ledgers if _sl['closing'] > 0]
                if len(_nonzero_stk) > 1:
                    for _sl in _nonzero_stk:
                        ca_lines.append((f"  {_sl['name']}", -_sl['closing']))
            # else: bs_closing_stock=0 → drop the Opening Stock row entirely (no phantom line)
        else:
            ca_lines.append((_nm, _val))

    # Rolled-up group totals
    cap_total   = TB.get('Capital Account', 0)
    loans_total = TB.get('Loans (Liability)', 0)
    cl_total    = TB.get('Current Liabilities', 0)
    fa_total    = TB.get('Fixed Assets', 0)

    # CA total: Replace Opening Stock with bs_closing_stock (matches Tally BS).
    # For no-stock companies (bs_closing_stock=0): subtract Opening Stock entirely
    # so the phantom OS row does not inflate the CA total.
    _ca_tb        = TB.get('Current Assets', 0)
    _open_s_in_ca = TB.get('Opening Stock', 0)
    _ca_has_os    = any(n == 'Opening Stock' for n, v in _ca_raw)
    if _ca_has_os:
        if bs_closing_stock > 0:
            # Stock company: swap OS → CS
            ca_total = _ca_tb - _open_s_in_ca + (-bs_closing_stock)
        else:
            # No-stock company: remove Opening Stock row entirely from CA total
            ca_total = _ca_tb - _open_s_in_ca
    else:
        ca_total = _ca_tb

    # Cache tb_children_of for ALL groups so build_bs can use any group dynamically
    tb_children_cache = {}
    for _n in TB_ORDER:
        if PARENT.get(_n,'') == 'Primary' or _n in {
            'Capital Account','Loans (Liability)','Current Liabilities',
            'Fixed Assets','Current Assets','Investments',
            'Misc. Expenses (ASSET)','Branch / Divisions','Suspense A/c',
        }:
            tb_children_cache[_n] = tb_children_of(_n)

    # P&L A/c (on Tally BS) — DIRECT computation, not balancing figure.
    # Tally's BS shows:  P&L A/c = Opening Balance + Current Period
    # P&L on BS (Tally convention):
    #   TB['Profit & Loss A/c'] = OPENING balance brought forward from prior years.
    #   TB sign: POSITIVE = CR balance = accumulated LOSS (shown on ASSET side of Tally BS as debit).
    #            NEGATIVE = DR balance = accumulated PROFIT (shown on LIABILITY side as credit).
    #   Current year net profit flows through P&L groups (Sales, Expenses etc.) — NOT into TB P&L group.
    #   So BS P&L = Opening (TB value) + Current year net profit.
    #
    # FIX: Tally stores an accumulated loss as a POSITIVE number in TB['Profit & Loss A/c']
    # because it lives on the asset side. We must NEGATE it so that:
    #   pnl_bf  > 0 → means accumulated PROFIT (adds to reserves correctly)
    #   pnl_bf  < 0 → means accumulated LOSS   (reduces reserves correctly)
    # Without this, a loss company shows a POSITIVE reserve → wrong BS.
    # P&L A/c TB sign convention (verified against live Tally data):
    # Tally uses STANDARD double-entry signs in the TB XML:
    #   DR balance (loss accumulated) → NEGATIVE value in TB dict
    #   CR balance (profit accumulated) → POSITIVE value in TB dict
    # So: TB['Profit & Loss A/c'] = -1327530 means accumulated LOSS of 1327530
    #     TB['Profit & Loss A/c'] = +500000  means accumulated PROFIT of 500000
    # We use the raw value directly — NO negation needed.
    # pnl_bf < 0 → loss → reduces reserves → BS liabilities decrease → correct
    # pnl_bf > 0 → profit → increases reserves → BS liabilities increase → correct
    _pnl_bf_raw = TB.get('Profit & Loss A/c', 0)
    pnl_bf      = _pnl_bf_raw    # use directly — DR=negative=loss, CR=positive=profit
    pnl_opening = pnl_bf
    pnl_current = net_profit     # already signed: profit=positive, loss=negative
    pnl_on_bs   = pnl_bf + net_profit  # total P&L on BS

    # Fixed asset helpers for Note 8
    dep_reserve = abs(TB.get('DEPRICIATION RESERVE', TB.get('Depreciation Reserve', 0)))
    fa_gross    = sum(abs(v) for n, v in fa_lines_all if 'depreci' not in n.lower())
    fa_net_bs   = abs(fa_total)   # Tally BS shows net (gross − dep) as group total

    # Equity share capital — SUM ALL matching ledgers (handles multi-shareholder Pvt Ltd,
    # LLPs, and partner capital accounts where each partner has a separate "XYZ Share Capital" ledger).
    # FIX: was breaking on first match → wrong for companies like Alpshrey with two share capital ledgers.
    equity_share = 0
    # Pass 1: TB-level keys that explicitly say "equity share capital"
    for k, v in TB.items():
        if 'equity share' in k.lower() or ('share capital' in k.lower() and 'equity' in k.lower()):
            equity_share += abs(v)
    if not equity_share:
        # Pass 2: cap_lines — sum ALL ledgers whose name contains 'share capital' or 'equity'
        # This catches "Alpesh Ranpariya Share Capital" + "Shreya Share Capital" style naming.
        for n, v in cap_lines:
            if 'share capital' in n.lower() or 'equity' in n.lower():
                equity_share += abs(v)
    if not equity_share:
        # Pass 3: if still zero, use full cap_total as equity (proprietorship / single-owner)
        # Only when cap_total is a CR balance (positive = liability side = genuine equity)
        _ct = TB.get('Capital Account', 0)
        if _ct > 0:
            equity_share = abs(_ct)

    reserves    = abs(TB.get('Reserves & Surplus', TB.get('Retained Earnings', 0)))
    debtors_net = abs(TB.get('Sundry Debtors', 0))
    income_tax  = abs(TB.get('Income Tax Exp.', TB.get('Income Tax', 0)))

    # ── PREVIOUS YEAR TB — for BS/PNL second column ───────────────────────────
    # Parses "Trial Balance PY" data in exactly the same way as current year.
    # All PY values suffixed with _py. Empty dict if PY fetch failed.
    TB_PY = {}
    TB_ORDER_PY = []
    py_tb_raw = data.get("Trial Balance PY", {})
    if isinstance(py_tb_raw, dict) and 'LINEERROR' not in py_tb_raw:
        try:
            py_tb_names = _find_key(py_tb_raw, 'DSPACCNAME') or []
            py_tb_info  = _find_key(py_tb_raw, 'DSPACCINFO') or []
            if not isinstance(py_tb_names, list): py_tb_names = [py_tb_names]
            if not isinstance(py_tb_info,  list): py_tb_info  = [py_tb_info]
            for n, info in zip(py_tb_names, py_tb_info):
                pname = n.get('DSPDISPNAME', '') if isinstance(n, dict) else ''
                if not pname: continue
                dr = fv(info.get('DSPCLDRAMT', {}).get('DSPCLDRAMTA') if isinstance(info, dict) else 0)
                cr = fv(info.get('DSPCLCRAMT', {}).get('DSPCLCRAMTA') if isinstance(info, dict) else 0)
                TB_PY[pname] = dr + cr
                TB_ORDER_PY.append(pname)
        except Exception:
            pass

    # PY stock: use Outstanding Bills PY closing balances (same logic as current year)
    _closing_stock_py = sum(
        abs(fv(item.get('CLOSINGBALANCE', 0)))
        for item in ob_py_list
        if isinstance(item, dict) and item.get('PARENT', '') == 'Stock-in-Hand'
    )
    _opening_stock_py = abs(TB_PY.get('Opening Stock', 0))

    # PY P&L figures
    sales_py      = TB_PY.get('Sales Accounts', 0)
    purch_py      = TB_PY.get('Purchase Accounts', 0)
    dir_inc_py    = TB_PY.get('Direct Incomes', 0)
    dir_exp_py    = TB_PY.get('Direct Expenses', 0)
    ind_inc_py    = TB_PY.get('Indirect Incomes', 0)
    ind_exp_py    = TB_PY.get('Indirect Expenses', 0)
    pnl_bf_py     = TB_PY.get('Profit & Loss A/c', 0)   # use directly — same sign convention as CY

    # Use Native PnL PY for closing stock if available (Type B companies)
    native_pnl_py = data.get("Native PnL PY", {})
    _native_cs_py = native_pnl_py.get("closing_stock") if isinstance(native_pnl_py, dict) else None
    if _native_cs_py and _native_cs_py > 0:
        _closing_stock_py = _native_cs_py

    # PY stock terms: same guard as current year.
    # If the current year company is trading-without-inventory or service,
    # the previous year is almost certainly the same type — zero PY stock terms too.
    # This prevents a stale _py_closing value (fetched from OB PY) from inflating PY GP.
    if _is_services or _is_trading_no_inventory:
        _closing_stock_py = 0.0
        _opening_stock_py = 0.0

    gross_profit_py = (sales_py + dir_inc_py + _closing_stock_py
                       + purch_py + (-_opening_stock_py) + dir_exp_py)
    net_profit_py   = gross_profit_py + ind_inc_py + ind_exp_py

    # PY BS group totals (for previous year column on BS)
    cap_total_py    = TB_PY.get('Capital Account', 0)
    loans_total_py  = TB_PY.get('Loans (Liability)', 0)
    cl_total_py     = TB_PY.get('Current Liabilities', 0)
    fa_total_py     = TB_PY.get('Fixed Assets', 0)
    _ca_py_tb       = TB_PY.get('Current Assets', 0)
    _os_py_in_ca    = TB_PY.get('Opening Stock', 0)
    _py_has_os_in_ca = _os_py_in_ca != 0
    if _closing_stock_py > 0:
        # PY stock company: swap OS → CS in CA total
        ca_total_py = _ca_py_tb - _os_py_in_ca + (-_closing_stock_py)
    elif _py_has_os_in_ca and (_is_services or _is_trading_no_inventory):
        # PY no-stock company: remove phantom Opening Stock from CA total
        ca_total_py = _ca_py_tb - _os_py_in_ca
    else:
        ca_total_py = _ca_py_tb
    pnl_on_bs_py    = pnl_bf_py + net_profit_py     # opening + current year profit

    return {
        # ── meta ──────────────────────────────────────────────────────────────
        'company':   company,
        'fy_start':  fy_start,
        'fy_end':    fy_end,
        'co_info':   co_info,
        # ── raw lookups ───────────────────────────────────────────────────────
        'TB':            TB,
        'TB_ORDER':      TB_ORDER,
        'PARENT':        PARENT,
        'get_primary':   get_primary,
        'direct_children': direct_children,
        'child_lines':   child_lines,
        # ── P&L values ────────────────────────────────────────────────────────
        'sales':         sales_grp,
        'purchases':     purch_grp,
        'open_s':        open_s_grp,
        'dir_inc':       dir_inc_grp,
        'dir_exp':       dir_exp_grp,
        'ind_inc':       ind_inc_grp,
        'ind_exp':       ind_exp_grp,
        'pnl_bf':        pnl_bf,
        'gross_profit':  gross_profit,
        'net_profit':    net_profit,
        # P&L child lines (dynamic — whatever Tally has)
        'sales_lines':    sales_lines,
        'purch_lines':    purch_lines,
        'dir_inc_lines':  dir_inc_lines,
        'dir_exp_lines':  dir_exp_lines,
        'ind_inc_lines':  ind_inc_lines,
        'ind_exp_lines':  ind_exp_lines,
        # ── stock ─────────────────────────────────────────────────────────────
        'closing_stock_abs': closing_stock_abs,
        'opening_stock_abs': opening_stock_abs,
        'bs_closing_stock':  bs_closing_stock,   # used on BS assets + P&L Cr side
        'stock_by_cat':      stock_by_cat,
        'stk_ledgers':          stk_ledgers,        # [{name,closing,opening}] current year
        'stk_ledgers_py':       stk_ledgers_py,     # [{name,closing,opening}] previous year
        'closing_stock_abs_py': closing_stock_abs_py,  # prev year closing stock
        # ── BS group totals (Tally rolled-up net values) ──────────────────────
        'tb_children_of_cache': tb_children_cache,
        'cap_total':    cap_total,
        'loans_total':  loans_total,
        'cl_total':     cl_total,
        'fa_total':     fa_total,
        'ca_total':     ca_total,
        # ── BS child lines (dynamic) ──────────────────────────────────────────
        'cap_lines':    cap_lines,
        'loans_lines':  loans_lines,
        'cl_lines':     cl_lines,
        'fa_lines_all': fa_lines_all,
        'ca_lines':     ca_lines,
        # ── P&L A/c (shown on BS) ─────────────────────────────────────────────
        'pnl_on_bs':    pnl_on_bs,
        'pnl_opening':  pnl_opening,
        'pnl_current':  pnl_current,
        # ── Previous Year (PY) data — for second column in BS / P&L ──────────
        'TB_PY':              TB_PY,
        'TB_ORDER_PY':        TB_ORDER_PY,
        'gross_profit_py':    gross_profit_py,
        'net_profit_py':      net_profit_py,
        'sales_py':           sales_py,
        'purch_py':           purch_py,
        'dir_inc_py':         dir_inc_py,
        'dir_exp_py':         dir_exp_py,
        'ind_inc_py':         ind_inc_py,
        'ind_exp_py':         ind_exp_py,
        'pnl_bf_py':          pnl_bf_py,
        'pnl_on_bs_py':       pnl_on_bs_py,
        'closing_stock_py':   _closing_stock_py,
        'opening_stock_py':   _opening_stock_py,
        'cap_total_py':       cap_total_py,
        'loans_total_py':     loans_total_py,
        'cl_total_py':        cl_total_py,
        'fa_total_py':        fa_total_py,
        'ca_total_py':        ca_total_py,
        # ── Fixed asset helpers ───────────────────────────────────────────────
        'dep_reserve':  dep_reserve,
        'fa_gross':     fa_gross,
        'fa_net_bs':    fa_net_bs,
        # ── commonly-used individual values (for notes) ───────────────────────
        'equity_share': equity_share,
        'reserves':     reserves,
        'debtors_net':  debtors_net,
        'income_tax':   income_tax,
        # outstanding
        'debtors_list':   debtors_list,
        'creditors_list': creditors_list,
        # kept for backward compat with Notes section
        'dtl':           abs(TB.get('Deffered Tax Liabilites', TB.get('Deferred Tax Liability', 0))),
        'bank_od':        abs(TB.get('Bank OD A/c', 0)),
        'secured_loans':  abs(TB.get('Secured Loans', TB.get('Secured Loans -', 0))),
        'unsecured_loans':abs(TB.get('Unsecured Loans', 0)),
        'sundry_cred':    abs(TB.get('Sundry Creditors', 0)),
        'outstanding_e':  abs(TB.get('OUTSTANDING EXPENSES', 0)),
        'stat_liab':      abs(TB.get('STATUTORY LIABILITIES', 0)),
        'duties_taxes':   abs(TB.get('Duties & Taxes', 0)),
        'provisions':     abs(TB.get('Provisions', 0)),
        'cash_hand':      abs(TB.get('Cash-in-Hand', TB.get('Cash in Hand', 0))),
        'bank_accts':     abs(TB.get('Bank Accounts', 0)),
        'loans_adv':      abs(TB.get('LOAN & ADVANCES', TB.get('Loan & Advances', 0))),
        'deposits':       abs(TB.get('Deposits (Asset)', 0)),
        'taxation_a':     abs(TB.get('Taxation', 0)),
        'other_ca':       abs(TB.get('Other Current Asset and Prepaid Exps.', 0)),
        'emda':           abs(TB.get('EMD A/C', 0)),
        'dbk':            abs(TB.get('DBK Receiveble  ( Export)', TB.get('DBK Receivable', 0))),
        'rodtep':         abs(TB.get('Export Incentives RODTEP Receiveble', TB.get('RODTEP Receivable', 0))),
        'diff_duty':      abs(TB.get('Diffrential Duty Recovreble Customs', 0)),
        'rodtep_lic':     abs(TB.get('RODTEP LICENCE IN HAND', 0)),
        'mfg_exp':        abs(TB.get('Manufacturing & Others Expenses', 0)),
        'power_fuel':     abs(TB.get('Power & Fuel Expenses', TB.get('Power & Fuel', 0))),
        'stores_exp':     abs(TB.get('Purchases-Stores and Consumables', 0)),
        'stores_cons':    abs(TB.get('Stores & Consumable Expenses.', 0)),
        'purch_rm':       abs(TB.get('Purchases-Raw Materials', 0)),
        'purch_job':      abs(TB.get('JOB PURCHASE', 0)),
        'purch_clg':      abs(TB.get('PURCHASES-CLG & FORWARDING EXPENSES', 0)),
        'emp_exp':        abs(TB.get('Payment to Employees', 0)),
        'finance_exp':    abs(TB.get('Financial Expenses', 0)),
        'admin_exp':      abs(TB.get('Administrative Exps', TB.get('Administrative Expenses', 0))),
        'factory_exp':    abs(TB.get('Factory Expenses', 0)),
        'legal_exp':      abs(TB.get('Legal and Prof. Exps', TB.get('Legal & Professional Expenses', 0))),
        'selling_exp':    abs(TB.get('Selling and Distribution Exps.', TB.get('Selling Expenses', 0))),
        'sales_domestic': abs(TB.get('SALES-DOMESTIC', TB.get('Sales-Domestic', 0))),
        'sales_export':   abs(TB.get('SALES-EXPORTS',  TB.get('Sales-Exports', 0))),
        'sales_jobwork':  abs(TB.get('SALES- JOB WORK',TB.get('Sales - Job Work', 0))),
    }

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — PDF STYLES & HELPERS
# ══════════════════════════════════════════════════════════════════════════════
C_DARK  = colors.HexColor("#1F3864")
C_MID   = colors.HexColor("#2E5496")
C_LIGHT = colors.HexColor("#DCE6F1")
C_BORD  = colors.HexColor("#8EA9C1")
C_GREY  = colors.HexColor("#666666")
PAGE_W  = 190 * mm

def mk_styles():
    S = getSampleStyleSheet()
    def ps(name, **kw):
        base = kw.pop('parent', S["Normal"])
        return ParagraphStyle(name, parent=base, **kw)
    return {
        "H1":   ps("H1",  fontSize=13, fontName="Helvetica-Bold", textColor=C_DARK,
                          alignment=TA_CENTER, spaceAfter=2, leading=16),
        "H2":   ps("H2",  fontSize=8,  fontName="Helvetica", textColor=C_GREY,
                          alignment=TA_CENTER, spaceAfter=3, leading=10),
        "NRM":  ps("NRM", fontSize=7.5,fontName="Helvetica", spaceAfter=1, leading=10),
        "IND":  ps("IND", fontSize=7.5,fontName="Helvetica", spaceAfter=1, leading=10,
                          leftIndent=10),
        "IND2": ps("IND2",fontSize=7,  fontName="Helvetica", spaceAfter=1, leading=9,
                          leftIndent=20, textColor=C_GREY),
        "BOLD": ps("BOLD",fontSize=7.5,fontName="Helvetica-Bold", spaceAfter=1, leading=10,
                          textColor=C_DARK),
        "AMT":  ps("AMT", fontSize=7.5,fontName="Helvetica", spaceAfter=1, leading=10,
                          alignment=TA_RIGHT),
        "AMTb": ps("AMTb",fontSize=7.5,fontName="Helvetica-Bold", spaceAfter=1, leading=10,
                          alignment=TA_RIGHT, textColor=C_DARK),
        "CTR":  ps("CTR", fontSize=7.5,fontName="Helvetica", spaceAfter=1, leading=10,
                          alignment=TA_CENTER),
        "SML":  ps("SML", fontSize=6.5,fontName="Helvetica", spaceAfter=1, leading=9,
                          textColor=C_GREY),
    }

def fmt_indian(n):
    """Format number in Indian numbering (crores/lakhs)."""
    n = abs(float(n))
    if n == 0: return "-"
    # split into integer and decimal
    integer = str(int(n))
    dec     = f"{n:.2f}".split('.')[1]
    if len(integer) <= 3:
        return f"{integer}.{dec}"
    last3 = integer[-3:]
    rest  = integer[:-3]
    parts = []
    while len(rest) > 2:
        parts.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest: parts.insert(0, rest)
    return ','.join(parts) + ',' + last3 + '.' + dec

def amt(v, blank_zero=False):
    if v is None or float(v) == 0:
        return "" if blank_zero else "-"
    v = float(v)
    return f"({fmt_indian(abs(v))})" if v < 0 else fmt_indian(v)

def P(txt, st="NRM", indent=0):
    if isinstance(st, int): indent, st = st, "NRM"
    if indent == 1: st = "IND"
    if indent == 2: st = "IND2"
    return Paragraph(txt, STYLES[st])

def A(v):  return Paragraph(amt(v),  STYLES["AMT"])
def AB(v): return Paragraph(amt(v),  STYLES["AMTb"])
def N(n):  return Paragraph(str(n),  STYLES["CTR"])

STYLES = {}  # populated in generate_pdf

def col_header(labels, cw, extra_style=None):
    ts = TableStyle([
        ("BACKGROUND",(0,0),(-1,0), C_DARK),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),7.5),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(0,-1),6),
        ("GRID",(0,0),(-1,-1),0.2,C_BORD),
    ])
    if extra_style:
        for s in extra_style: ts.add(*s)
    return Table([[Paragraph(f"<b>{l}</b>", STYLES["BOLD"]) for l in labels]],
                 colWidths=cw, style=ts)

def note_hdr(num, title, story):
    story.append(Spacer(1, 3*mm))
    story.append(Table(
        [[Paragraph(f"<b>Note {num} :  {title}</b>", STYLES["BOLD"])]],
        colWidths=[PAGE_W],
        style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), C_DARK),
            ("TEXTCOLOR",(0,0),(-1,-1), colors.white),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),6),
        ])))

def data_tbl(rows, cw, story, total_last=True, zebra=True):
    if not rows: return
    ts = [
        ("FONTSIZE",(0,0),(-1,-1),7.5),
        ("TOPPADDING",(0,0),(-1,-1),2.5),("BOTTOMPADDING",(0,0),(-1,-1),2.5),
        ("LEFTPADDING",(0,0),(0,-1),5),("RIGHTPADDING",(-1,0),(-1,-1),4),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("GRID",(0,0),(-1,-1),0.2,C_BORD),
    ]
    if zebra:
        for i in range(1, len(rows), 2):
            ts.append(("BACKGROUND",(0,i),(-1,i), C_LIGHT))
    if total_last:
        ts += [
            ("BACKGROUND",(0,-1),(-1,-1), C_MID),
            ("TEXTCOLOR",(0,-1),(-1,-1), colors.white),
            ("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
        ]
    story.append(Table(rows, colWidths=cw, style=TableStyle(ts)))

def bs_tbl(rows, cw, story):
    """Balance sheet main table with smart row colouring."""
    ts = [
        ("FONTSIZE",(0,0),(-1,-1),7.5),
        ("TOPPADDING",(0,0),(-1,-1),2.5),("BOTTOMPADDING",(0,0),(-1,-1),2.5),
        ("LEFTPADDING",(0,0),(0,-1),4),("RIGHTPADDING",(-1,0),(-1,-1),4),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("GRID",(0,0),(-1,-1),0.2,C_BORD),
    ]
    for i, row in enumerate(rows):
        txt = row[0].text if hasattr(row[0], 'text') else ''
        if any(k in txt for k in ['TOTAL EQUITY','TOTAL ASSETS']):
            ts += [("BACKGROUND",(0,i),(-1,i),C_DARK),
                   ("TEXTCOLOR",(0,i),(-1,i),colors.white),
                   ("FONTNAME",(0,i),(-1,i),"Helvetica-Bold")]
        elif txt.strip().startswith('Total '):
            ts += [("BACKGROUND",(0,i),(-1,i),C_MID),
                   ("TEXTCOLOR",(0,i),(-1,i),colors.white),
                   ("FONTNAME",(0,i),(-1,i),"Helvetica-Bold")]
        elif txt.strip().startswith(('1.','2.','3.','4.','5.')):
            ts += [("BACKGROUND",(0,i),(-1,i),colors.HexColor("#EEF4FB")),
                   ("FONTNAME",(0,i),(0,i),"Helvetica-Bold")]
        elif txt in ('EQUITY AND LIABILITIES','ASSETS'):
            ts += [("BACKGROUND",(0,i),(-1,i),colors.HexColor("#F0F0F0")),
                   ("FONTNAME",(0,i),(-1,i),"Helvetica-Bold")]
        elif i % 2 == 0:
            ts += [("BACKGROUND",(0,i),(-1,i),C_LIGHT)]
    story.append(Table(rows, colWidths=cw, style=TableStyle(ts)))


# ══════════════════════════════════════════════════════════════════════════════
# TRIAL BALANCE  — exact Tally format (group rows gold+bold, ledgers indented)
# Uses TB_ORDER from parse_data for correct positional ordering
# ══════════════════════════════════════════════════════════════════════════════
def build_tb(D, story):
    CW = [PAGE_W * 0.52, PAGE_W * 0.24, PAGE_W * 0.24]

    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(D['company'], STYLES["H1"]))
    story.append(Paragraph("TRIAL BALANCE",
        ParagraphStyle("T", parent=STYLES["H1"], fontSize=11, spaceAfter=1)))
    story.append(Paragraph(f"For the period  {D['fy_start']}  to  {D['fy_end']}", STYLES["H2"]))
    story.append(Paragraph("(All amounts in Indian Rupees)", STYLES["H2"]))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_DARK, spaceAfter=3))
    story.append(col_header(["Particulars", "Debit", "Credit"], CW))

    TB       = D['TB']
    TB_ORDER = D['TB_ORDER']
    PARENT   = D['PARENT']

    # Identify which names are group rows (have children in TB_ORDER)
    # A name is a group if it has at least one child row immediately after it
    grp_set = set()
    for i, name in enumerate(TB_ORDER):
        if i + 1 < len(TB_ORDER):
            next_name = TB_ORDER[i + 1]
            # next row is a child if its primary = current row's primary, or parent = this name
            if PARENT.get(next_name, '') == name or name in {
                'Capital Account','Loans (Liability)','Current Liabilities',
                'Fixed Assets','Current Assets','Sales Accounts',
                'Purchase Accounts','Direct Expenses','Indirect Incomes','Indirect Expenses',
            }:
                grp_set.add(name)
        # Also: any name that is a known primary group
        if PARENT.get(name, '') == 'Primary' or name in {
            'Capital Account','Loans (Liability)','Current Liabilities',
            'Fixed Assets','Current Assets','Sales Accounts',
            'Purchase Accounts','Direct Expenses','Indirect Incomes',
            'Indirect Expenses','Direct Incomes','Investments',
        }:
            grp_set.add(name)

    rows = []
    grand_dr = 0.0
    grand_cr = 0.0

    GOLD_C  = colors.HexColor("#F5C518")
    GOLD_T  = colors.HexColor("#5A3A00")

    grp_indices  = []   # indices of group rows (for gold background)
    row_types    = []   # 'grp' or 'led' for each row

    for name in TB_ORDER:
        net = TB.get(name, 0)
        dr_val = abs(net) if net < 0 else 0
        cr_val = abs(net) if net > 0 else 0
        grand_dr += dr_val
        grand_cr += cr_val

        dr_txt = fmt_indian(dr_val) if dr_val > 0 else "-"
        cr_txt = fmt_indian(cr_val) if cr_val > 0 else "-"
        is_grp = name in grp_set

        if is_grp:
            rows.append([
                Paragraph(f"<b>{name}</b>",
                    ParagraphStyle("G", parent=STYLES["BOLD"], textColor=GOLD_T)),
                Paragraph(dr_txt,
                    ParagraphStyle("Gd", parent=STYLES["AMTb"], textColor=GOLD_T)),
                Paragraph(cr_txt,
                    ParagraphStyle("Gc", parent=STYLES["AMTb"], textColor=GOLD_T)),
            ])
            row_types.append('grp')
        else:
            rows.append([
                P(name, "IND"),
                Paragraph(dr_txt, STYLES["AMT"]),
                Paragraph(cr_txt, STYLES["AMT"]),
            ])
            row_types.append('led')

    # Grand total
    rows.append([
        P("<b>Grand Total</b>", "BOLD"),
        Paragraph(fmt_indian(grand_dr), STYLES["AMTb"]),
        Paragraph(fmt_indian(grand_cr), STYLES["AMTb"]),
    ])
    row_types.append('tot')

    ts = [
        ("FONTSIZE",      (0, 0), (-1, -1), 7.5),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING",   (0, 0), (0, -1),  4),
        ("RIGHTPADDING",  (-1,0), (-1,-1),  4),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.2, C_BORD),
        # Grand total row
        ("BACKGROUND", (0, -1), (-1, -1), C_DARK),
        ("TEXTCOLOR",  (0, -1), (-1, -1), colors.white),
        ("FONTNAME",   (0, -1), (-1, -1), "Helvetica-Bold"),
    ]
    for i, rt in enumerate(row_types[:-1]):   # exclude grand total
        if rt == 'grp':
            ts.append(("BACKGROUND", (0, i), (-1, i), GOLD_C))
        elif i % 2 == 0:
            ts.append(("BACKGROUND", (0, i), (-1, i), C_LIGHT))

    story.append(Table(rows, colWidths=CW, style=TableStyle(ts)))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f"Debit Total: {fmt_indian(grand_dr)}   |   Credit Total: {fmt_indian(grand_cr)}",
        STYLES["SML"]))

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — BALANCE SHEET  (fully dynamic — mirrors Tally BS exactly)
# ══════════════════════════════════════════════════════════════════════════════
def build_bs(D, story):
    """
    Balance Sheet — exact Tally logic:
    Tally puts each primary group on L or A side purely by its NET SIGN:
      • Positive net (CR balance) → Liabilities side
      • Negative net (DR balance) → Assets side
    This handles ALL cases:
      - Current Assets with CR net → L side (e.g. Venkateshwara Traders)
      - P&L A/c with DR net (accumulated loss) → A side
      - Investments → A side (DR)
      - Any custom primary group → correct side by sign
    """
    TB   = D['TB']
    CW   = [PAGE_W * 0.46, PAGE_W * 0.06, PAGE_W * 0.24, PAGE_W * 0.24]
    fy_e = D['fy_end']
    fy_p = f"{int(fy_e[:4])-1}{fy_e[4:]}"

    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(D['company'], STYLES["H1"]))
    story.append(Paragraph("BALANCE SHEET",
        ParagraphStyle("T", parent=STYLES["H1"], fontSize=11, spaceAfter=1)))
    story.append(Paragraph(f"As at  {fy_e}", STYLES["H2"]))
    story.append(Paragraph("(All amounts in Indian Rupees unless otherwise stated)", STYLES["H2"]))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_DARK, spaceAfter=3))
    story.append(col_header(["Particulars", "Note", f"As at {fy_e}", f"As at {fy_p}"], CW))

    # ── Pre-compute group nets ────────────────────────────────────────────────
    # All known BS primary groups (P&L computed separately as balancing figure)
    PL_GROUPS = {'Sales Accounts','Purchase Accounts','Direct Expenses',
                 'Direct Incomes','Indirect Expenses','Indirect Incomes'}
    BS_PRIMARY = {
        'Capital Account','Loans (Liability)','Current Liabilities',
        'Fixed Assets','Current Assets','Investments',
        'Misc. Expenses (ASSET)','Branch / Divisions','Suspense A/c',
    }

    group_net = {}
    PARENT = D['PARENT']
    for name in D['TB_ORDER']:
        if (PARENT.get(name,'') == 'Primary' or name in BS_PRIMARY) and name not in PL_GROUPS:
            group_net[name] = TB.get(name, 0)

    # Override CA with stock-adjusted value
    if 'Current Assets' in group_net:
        group_net['Current Assets'] = D['ca_total']

    # P&L balancing figure
    pnl_opening = D['pnl_opening']
    pnl_current = D['pnl_current']
    pnl_net     = D.get('pnl_on_bs', pnl_opening + pnl_current)

    # ── Map group → child lines ───────────────────────────────────────────────
    GROUP_CHILDREN = {
        'Capital Account':     D.get('cap_lines',    []),
        'Loans (Liability)':   D.get('loans_lines',  []),
        'Current Liabilities': D.get('cl_lines',     []),
        'Fixed Assets':        D.get('fa_lines_all', []),
        'Current Assets':      D.get('ca_lines',     []),
    }
    for name in group_net:
        if name not in GROUP_CHILDREN:
            GROUP_CHILDREN[name] = D.get('tb_children_of_cache', {}).get(name, [])

    def blank():
        return [P("", "NRM"), N(""), A(None), A(None)]

    def group_hdr_row(label):
        return [P(label, "NRM"), N(""), A(None), A(None)]

    def group_total_row(label, val, val_py=None):
        return [P(label, "BOLD"), N(""), AB(val), AB(val_py)]

    def child_rows_for(group_name, is_asset=False):
        """
        Emit child rows for a group, handling all sign scenarios:

        LIABILITIES side (is_asset=False):
          - Normal CR child (positive): show as positive
          - DR child (negative): show as (-)value — e.g. Duties & Taxes DR, PnL(Dr/Cr) DR
          - Zero: skip

        ASSETS side (is_asset=True):
          - Normal DR child (negative in TB): show as positive (abs)
          - CR child (positive in TB, unusual): show as (-)value — deduction from assets
          - Zero: skip
        """
        lines = GROUP_CHILDREN.get(group_name, [])
        out = []
        for name, val in lines:
            if val == 0:
                continue
            if is_asset:
                # Asset side: DR values are normal (show positive), CR values are deductions (show negative)
                display_val = abs(val) if val < 0 else -abs(val)
            else:
                # Liability side: CR values are normal (show positive), DR values show as negative
                display_val = val  # keep sign: positive=CR=normal, negative=DR=shown in brackets
            out.append([P(f"    {name}", "IND"), N(""), A(display_val), A(None)])
        return out

    def fa_child_rows():
        """Fixed Assets: dep reserve shown as negative, assets as positive."""
        out = []
        dep = D['dep_reserve']
        if dep > 0:
            out.append([P("    DEPRECIATION RESERVE", "IND"), N(""), A(-dep), A(None)])
        for name, val in D.get('fa_lines_all', []):
            if name.upper() in ('DEPRICIATION RESERVE','DEPRECIATION RESERVE'):
                continue
            out.append([P(f"    {name}", "IND"), N(""), A(abs(val)), A(None)])
        return out

    # ── Classify each group into L or A side by net sign ─────────────────────
    # Rule: CR net (positive) → L side | DR net (negative) → A side | Zero → skip
    # This handles all scenarios: accumulated loss in Capital (→ A side),
    # CA with CR balance (→ L side), Investments DR (→ A side), etc.
    L_groups = []
    A_groups = []

    for name, net in group_net.items():
        if net == 0:
            continue  # skip zero-balance groups entirely
        if net > 0:
            L_groups.append((name, net))
        else:
            A_groups.append((name, abs(net)))

    # P&L A/c: classify by sign of balancing figure; skip if zero
    if abs(pnl_net) > 0.01:
        if pnl_net >= 0:
            L_groups.append(('Profit & Loss A/c', pnl_net))
        else:
            A_groups.append(('Profit & Loss A/c', abs(pnl_net)))

    # ── PY group net lookup ───────────────────────────────────────────────────
    TB_PY = D.get('TB_PY', {})
    def _py_group(gname_):
        """Return PY absolute value for a BS group; 0 if not available."""
        if not TB_PY: return None
        if gname_ == 'Current Assets':
            v = D.get('ca_total_py', TB_PY.get('Current Assets', 0))
        elif gname_ == 'Profit & Loss A/c':
            v = D.get('pnl_on_bs_py', 0)
        else:
            v = TB_PY.get(gname_, 0)
        return abs(v) if v else None

    # ══════════════════════════════════════════════════════════════════════════
    # BUILD ROWS
    # ══════════════════════════════════════════════════════════════════════════
    rows = []
    total_l    = 0.0
    total_a    = 0.0
    total_l_py = 0.0
    total_a_py = 0.0
    l_num = 0
    a_num = 0

    # ── LIABILITIES SIDE ─────────────────────────────────────────────────────
    rows.append([P("EQUITY AND LIABILITIES", "BOLD"), N(""), A(None), A(None)])
    for gname, gabs in L_groups:
        l_num += 1
        label_letter = chr(64 + l_num)   # A, B, C …
        total_l += gabs
        gabs_py = _py_group(gname)
        if gabs_py: total_l_py += gabs_py

        if gname == 'Profit & Loss A/c':
            pnl_on_bs_py = D.get('pnl_on_bs_py', None)
            rows.append(group_hdr_row(f"{l_num}.  Profit & Loss A/c"))
            rows.append([P("    Opening Balance", "IND"), N(""), A(pnl_opening), A(None)])
            rows.append([P("    Current Period",  "IND"), N(""), A(pnl_current), A(None)])
            rows.append(group_total_row(
                f"Total Profit & Loss A/c  ({label_letter})", pnl_net, pnl_on_bs_py))
        else:
            rows.append(group_hdr_row(f"{l_num}.  {gname}"))
            if gname == 'Fixed Assets':
                rows += fa_child_rows()
            else:
                rows += child_rows_for(gname)
            rows.append(group_total_row(
                f"Total {gname}  ({label_letter})", gabs, gabs_py))

    letters_l = '+'.join(chr(64+i) for i in range(1, l_num+1))
    rows.append(group_total_row(f"TOTAL  ({letters_l})", total_l,
                                total_l_py if total_l_py else None))
    rows.append(blank())

    # ── ASSETS SIDE ──────────────────────────────────────────────────────────
    rows.append([P("ASSETS", "BOLD"), N(""), A(None), A(None)])
    for gname, gabs in A_groups:
        a_num += 1
        label_letter = chr(64 + a_num)
        total_a += gabs
        gabs_py = _py_group(gname)
        if gabs_py: total_a_py += gabs_py

        if gname == 'Profit & Loss A/c':
            pnl_on_bs_py = D.get('pnl_on_bs_py', None)
            rows.append(group_hdr_row(f"{a_num}.  Profit & Loss A/c"))
            rows.append([P("    Opening Balance", "IND"), N(""), A(abs(pnl_opening)), A(None)])
            rows.append([P("    Current Period",  "IND"), N(""), A(abs(pnl_current)), A(None)])
            rows.append(group_total_row(
                f"Total Profit & Loss A/c  ({label_letter})", abs(pnl_net),
                abs(pnl_on_bs_py) if pnl_on_bs_py else None))
        elif gname == 'Fixed Assets':
            rows.append(group_hdr_row(f"{a_num}.  Fixed Assets"))
            rows += fa_child_rows()
            rows.append(group_total_row(
                f"Total Fixed Assets  ({label_letter})", gabs, gabs_py))
        else:
            rows.append(group_hdr_row(f"{a_num}.  {gname}"))
            rows += child_rows_for(gname, is_asset=True)
            rows.append(group_total_row(
                f"Total {gname}  ({label_letter})", gabs, gabs_py))

    letters_a = '+'.join(chr(64+i) for i in range(1, a_num+1))
    rows.append(group_total_row(f"TOTAL  ({letters_a})", total_a,
                                total_a_py if total_a_py else None))

    # ── Render with smart colouring ──────────────────────────────────────────
    ts = [
        ("FONTSIZE",      (0,0), (-1,-1), 7.5),
        ("TOPPADDING",    (0,0), (-1,-1), 2.5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2.5),
        ("LEFTPADDING",   (0,0), (0,-1),  4),
        ("RIGHTPADDING",  (-1,0),(-1,-1), 4),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("GRID",          (0,0), (-1,-1), 0.2, C_BORD),
    ]
    for i, r in enumerate(rows):
        txt = r[0].text if hasattr(r[0], 'text') else ''
        if txt.startswith('TOTAL  (') or txt in ('EQUITY AND LIABILITIES','ASSETS'):
            ts += [("BACKGROUND",(0,i),(-1,i),C_DARK),
                   ("TEXTCOLOR", (0,i),(-1,i),colors.white),
                   ("FONTNAME",  (0,i),(-1,i),"Helvetica-Bold")]
        elif txt.startswith('Total '):
            ts += [("BACKGROUND",(0,i),(-1,i),C_MID),
                   ("TEXTCOLOR", (0,i),(-1,i),colors.white),
                   ("FONTNAME",  (0,i),(-1,i),"Helvetica-Bold")]
        elif txt.strip()[:3].rstrip('.').isdigit():
            ts += [("BACKGROUND",(0,i),(-1,i),colors.HexColor("#EEF4FB")),
                   ("FONTNAME",  (0,i),(0,i), "Helvetica-Bold")]
        elif i % 2 == 0:
            ts += [("BACKGROUND",(0,i),(-1,i),C_LIGHT)]

    story.append(Table(rows, colWidths=CW, style=TableStyle(ts)))

    # ── Signatory ─────────────────────────────────────────────────────────────
    story.append(Spacer(1, 5*mm))
    story.append(Table([[
        Paragraph("For and on behalf of the Board of Directors", STYLES["NRM"]),
        Paragraph("", STYLES["NRM"]),
        Paragraph("As per our Report of even date", STYLES["NRM"]),
    ]], colWidths=[PAGE_W*0.40, PAGE_W*0.20, PAGE_W*0.40],
    style=TableStyle([("FONTSIZE",(0,0),(-1,-1),7.5)])))
    story.append(Spacer(1, 8*mm))
    story.append(Table([[
        Paragraph("_______________________<br/>Director", STYLES["NRM"]),
        Paragraph("", STYLES["NRM"]),
        Paragraph("_______________________<br/>Chartered Accountant", STYLES["NRM"]),
    ]], colWidths=[PAGE_W*0.40, PAGE_W*0.20, PAGE_W*0.40],
    style=TableStyle([("FONTSIZE",(0,0),(-1,-1),7.5)])))
    story.append(Spacer(1, 2*mm))
    state = D['co_info'].get('STATENAME', 'India')
    story.append(Paragraph(f"Place: {state}   |   Date: {fy_e}", STYLES["SML"]))

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — PROFIT & LOSS  (fully dynamic — mirrors Tally P&L exactly)
# ══════════════════════════════════════════════════════════════════════════════
def build_pnl(D, story):
    TB   = D['TB']
    # Two-column T-format: Left = Dr side, Right = Cr side
    # Each side: [description, sub-amount, total-amount]
    # We lay it out as a SINGLE table with 6 columns:
    #   [L-desc, L-sub, L-total, R-desc, R-sub, R-total]
    CW6 = [PAGE_W*0.26, PAGE_W*0.10, PAGE_W*0.14,
           PAGE_W*0.26, PAGE_W*0.10, PAGE_W*0.14]

    fy_e = D['fy_end']
    fy_p = f"{int(fy_e[:4])-1}{fy_e[4:]}"

    story.append(PageBreak())
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(D['company'], STYLES["H1"]))
    story.append(Paragraph("PROFIT & LOSS ACCOUNT",
        ParagraphStyle("T", parent=STYLES["H1"], fontSize=11, spaceAfter=1)))
    story.append(Paragraph(f"1-Apr-{int(fy_e[:4])-1} to 31-Mar-{fy_e[:4]}", STYLES["H2"]))
    story.append(Paragraph("(All amounts in Indian Rupees unless otherwise stated)", STYLES["H2"]))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_DARK, spaceAfter=3))

    # ── Column header row ─────────────────────────────────────────────────────
    def ch(txt): return Paragraph(f"<b>{txt}</b>", STYLES["BOLD"])
    story.append(Table(
        [[ch("Particulars"), ch("Amount ₹"), ch(f"Total ₹"),
          ch("Particulars"), ch("Amount ₹"), ch(f"Total ₹")]],
        colWidths=CW6,
        style=TableStyle([
            ("BACKGROUND", (0,0), (-1,0), C_DARK),
            ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
            ("FONTSIZE",   (0,0), (-1,0), 7.5),
            ("ALIGN",      (1,0), (-1,0), "RIGHT"),
            ("TOPPADDING", (0,0), (-1,0), 4),
            ("BOTTOMPADDING",(0,0),(-1,0),4),
            ("LEFTPADDING",(0,0),(0,0),5),
            ("LEFTPADDING",(3,0),(3,0),5),
            ("GRID",       (0,0), (-1,0), 0.2, C_BORD),
            ("LINEAFTER",  (2,0), (2,0), 1.5, C_DARK),   # vertical divider
        ])))

    # ── Build Left (Dr) and Right (Cr) rows independently ─────────────────────
    # Each entry: (label, sub_val, total_val, is_header)
    # sub_val shown in col 1, total_val in col 2; None = blank cell
    left_rows  = []   # Dr side
    right_rows = []   # Cr side

    def L(lbl, sub=None, tot=None, hdr=False, indent=False):
        left_rows.append((lbl, sub, tot, hdr, indent))

    def R(lbl, sub=None, tot=None, hdr=False, indent=False):
        right_rows.append((lbl, sub, tot, hdr, indent))

    def Lblank(): left_rows.append(("", None, None, False, False))
    def Rblank(): right_rows.append(("", None, None, False, False))

    # ── Sign-normalisation helpers ─────────────────────────────────────────────
    # Tally TB convention:
    #   Expense / Asset groups  → DR  → stored as NEGATIVE in TB dict
    #   Income  / Liability groups → CR → stored as POSITIVE in TB dict
    #
    # For P&L display every amount is shown as POSITIVE (absolute).
    # "Deduction" lines (e.g. purchase-returns, income-credit-notes) are shown
    # in brackets — represented here as negative values passed to amt().
    #
    # dr_display(val): converts a DR-natured TB child value to display sign.
    #   Normal DR child (val < 0) → positive display (expense).
    #   Unusual CR child (val > 0) → negative display (deduction/return).
    def dr_display(val):
        # DR children (negative in TB) are normal expenses → show positive
        # CR children (positive in TB) are returns/deductions → show in brackets (negative)
        return abs(val) if val < 0 else -abs(val)

    # cr_display(val): converts a CR-natured TB child value to display sign.
    #   Normal CR child (val > 0) → positive display (income).
    #   Unusual DR child (val < 0) → negative display (deduction from income).
    def cr_display(val):
        return abs(val) if val > 0 else -abs(val)

    # ── LEFT SIDE (Dr / Expenditure) ──────────────────────────────────────────
    open_abs    = D['opening_stock_abs']
    pur_abs     = abs(D['purchases'])
    dir_abs     = abs(D['dir_exp'])
    gp          = D['gross_profit']
    net_profit  = D['net_profit']
    ind_exp_abs = abs(D['ind_exp'])

    # Opening Stock
    if open_abs > 0:
        L("Opening Stock", tot=open_abs, hdr=True)
        _stk_names = [(sl['name'], sl['opening'])
                      for sl in D.get('stk_ledgers', []) if sl.get('opening', 0) > 0]
        if not _stk_names:
            L("Stock in Hand", sub=open_abs, indent=True)
        elif len(_stk_names) == 1:
            L(_stk_names[0][0], sub=open_abs, indent=True)
        else:
            _total_raw = sum(v for _, v in _stk_names)
            for _nm, _raw in _stk_names:
                _share = open_abs * (_raw / _total_raw) if _total_raw > 0 else 0
                L(_nm, sub=_share, indent=True)

    # Purchase Accounts
    if pur_abs > 0:
        L("Purchase Accounts", tot=pur_abs, hdr=True)
        for name, val in sorted(D['purch_lines'], key=lambda x: abs(x[1]), reverse=True):
            if val != 0:
                # Purchase sub-lines are DR (negative in TB) → dr_display normalises
                L(name, sub=dr_display(val), indent=True)

    # Direct Expenses
    if dir_abs > 0:
        L("Direct Expenses", tot=dir_abs, hdr=True)
        for name, val in sorted(D['dir_exp_lines'], key=lambda x: abs(x[1]), reverse=True):
            if val != 0:
                L(name, sub=dr_display(val), indent=True)

    # Gross Profit / Loss carried over
    if gp >= 0:
        L("Gross Profit  c/o", tot=gp, hdr=True)
    else:
        L("Gross Loss  c/o", tot=abs(gp), hdr=True)
    Lblank()

    # Indirect Expenses
    if ind_exp_abs > 0:
        L("Indirect Expenses", tot=ind_exp_abs, hdr=True)
        for name, val in sorted(D['ind_exp_lines'], key=lambda x: abs(x[1]), reverse=True):
            if val != 0:
                # Indirect expense sub-lines are DR (negative in TB)
                L(name, sub=dr_display(val), indent=True)

    if net_profit >= 0:
        L("Net Profit  c/o", tot=net_profit, hdr=True)
    else:
        L("Net Loss  c/o", tot=abs(net_profit), hdr=True)
    Lblank()

    # ── RIGHT SIDE (Cr / Income) ───────────────────────────────────────────────
    sales_abs = abs(D['sales'])

    # Sales Accounts
    if sales_abs > 0:
        R("Sales Accounts", tot=sales_abs, hdr=True)
        for name, val in sorted(D['sales_lines'], key=lambda x: abs(x[1]), reverse=True):
            if val != 0:
                # Sales sub-lines are CR (positive in TB)
                R(name, sub=cr_display(val), indent=True)

    # Direct Incomes (affects Gross Profit)
    dir_inc_abs = abs(D.get('dir_inc', 0))
    if dir_inc_abs > 0:
        R("Direct Incomes", tot=dir_inc_abs, hdr=True)
        for name, val in sorted(D.get('dir_inc_lines', []), key=lambda x: abs(x[1]), reverse=True):
            if val != 0:
                R(name, sub=cr_display(val), indent=True)

    # Closing Stock
    close_abs = D.get('bs_closing_stock', D['closing_stock_abs'])
    if close_abs > 0:
        R("Closing Stock", tot=close_abs, hdr=True)
        # Use current year stk_ledgers CLOSING values (= stock as at fy_end)
        # These are the per-ledger breakdown of bs_closing_stock
        _close_items = [(sl['name'], sl['closing'])
                        for sl in D.get('stk_ledgers', []) if sl.get('closing', 0) > 0]
        if not _close_items:
            R("Stock in Hand", sub=close_abs, indent=True)
        elif len(_close_items) == 1:
            R(_close_items[0][0], sub=close_abs, indent=True)
        else:
            for _nm, _val in _close_items:
                R(_nm, sub=abs(_val), indent=True)

    Rblank()

    # Gross Profit / Loss brought forward
    if gp >= 0:
        R("Gross Profit  b/f", tot=gp, hdr=True)
    else:
        R("Gross Loss  b/f", tot=abs(gp), hdr=True)
    Rblank()

    # Indirect Incomes
    ind_inc_abs = abs(D['ind_inc'])
    if ind_inc_abs > 0:
        R("Indirect Incomes", tot=ind_inc_abs, hdr=True)
        for name, val in sorted(D['ind_inc_lines'], key=lambda x: abs(x[1]), reverse=True):
            if val != 0:
                # Indirect income sub-lines are CR (positive in TB)
                R(name, sub=cr_display(val), indent=True)

    if net_profit < 0:
        R("Net Loss  b/f", tot=abs(net_profit), hdr=True)

    Rblank()

    # ── Pad shorter side with blanks ──────────────────────────────────────────
    while len(left_rows) < len(right_rows):  Lblank()
    while len(right_rows) < len(left_rows):  Rblank()

    # ── Grand total row ────────────────────────────────────────────────────────
    # Tally T-format: both sides of the P&L must balance.
    # Dr side  = Opening Stock + Purchases + Direct Exp + Gross Profit c/o
    #            + Indirect Exp + Net Profit c/o
    # Cr side  = Sales + Closing Stock + Gross Profit b/f
    #            + Indirect Inc + Net Loss b/f  (if any)
    # The balancing total is Sales + Closing Stock + Indirect Incomes
    # (Net Profit is already embedded as the difference; adding it explicitly
    #  would double-count it on one side.)
    pnl_total = sales_abs + dir_inc_abs + close_abs + ind_inc_abs
    L("Total", tot=pnl_total, hdr=True)
    R("Total", tot=pnl_total, hdr=True)

    # ── Render combined table ──────────────────────────────────────────────────
    def fmt_cell(val):
        if val is None: return Paragraph("", STYLES["AMT"])
        return Paragraph(amt(val), STYLES["AMT"])

    def fmt_cell_b(val):
        if val is None: return Paragraph("", STYLES["AMTb"])
        return Paragraph(amt(val), STYLES["AMTb"])

    def pnl_p(lbl, is_hdr, indent):
        if is_hdr:  return Paragraph(f"<b>{lbl}</b>", STYLES["BOLD"])
        if indent:  return Paragraph(lbl, STYLES["IND"])
        return Paragraph(lbl, STYLES["NRM"])

    tbl_rows = []
    ts = [
        ("FONTSIZE",     (0,0), (-1,-1), 7.5),
        ("TOPPADDING",   (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0), (-1,-1), 2),
        ("LEFTPADDING",  (0,0), (0,-1),  4),
        ("LEFTPADDING",  (3,0), (3,-1),  4),
        ("RIGHTPADDING", (2,0), (2,-1),  4),
        ("RIGHTPADDING", (5,0), (5,-1),  4),
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
        ("GRID",         (0,0), (-1,-1), 0.2, C_BORD),
        ("LINEAFTER",    (2,0), (2,-1),  1.5, C_DARK),   # vertical divider between L and R
        ("ALIGN",        (1,0), (2,-1),  "RIGHT"),
        ("ALIGN",        (4,0), (5,-1),  "RIGHT"),
    ]

    for i, (lrow, rrow) in enumerate(zip(left_rows, right_rows)):
        ll, ls, lt, lh, li = lrow
        rl, rs, rt, rh, ri = rrow

        lp = pnl_p(ll, lh, li)
        rp = pnl_p(rl, rh, ri)
        lsub = fmt_cell(ls);   ltot = (fmt_cell_b if lh else fmt_cell)(lt)
        rsub = fmt_cell(rs);   rtot = (fmt_cell_b if rh else fmt_cell)(rt)

        tbl_rows.append([lp, lsub, ltot, rp, rsub, rtot])

        # Colour heading rows
        if lh:
            ts += [("BACKGROUND", (0,i), (2,i), colors.HexColor("#EEF4FB")),
                   ("FONTNAME",   (0,i), (2,i), "Helvetica-Bold")]
        if rh:
            ts += [("BACKGROUND", (3,i), (5,i), colors.HexColor("#EEF4FB")),
                   ("FONTNAME",   (3,i), (5,i), "Helvetica-Bold")]

        # Grand Total rows
        if ll == "Total" or rl == "Total":
            ts += [("BACKGROUND", (0,i), (-1,i), C_DARK),
                   ("TEXTCOLOR",  (0,i), (-1,i), colors.white),
                   ("FONTNAME",   (0,i), (-1,i), "Helvetica-Bold")]

    story.append(Table(tbl_rows, colWidths=CW6, style=TableStyle(ts)))

    # ── Schedule III Summary (for statutory filing) ────────────────────────────
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(
        "<b>Schedule III — Statement of Profit and Loss  (Summary Format for Statutory Compliance)</b>",
        STYLES["BOLD"]))
    story.append(Spacer(1, 2*mm))

    CW4 = [PAGE_W*0.50, PAGE_W*0.10, PAGE_W*0.20, PAGE_W*0.20]
    story.append(col_header(["Particulars","Note",f"Year ended {fy_e}",f"Year ended {fy_p}"], CW4))

    s3rows = []
    sales_tot = abs(D['sales'])
    ind_inc_t = abs(D['ind_inc'])
    tot_inc   = sales_tot + ind_inc_t

    # PY counterparts (None if PY TB not available — renders as "-")
    sales_tot_py = D.get('sales_py', None)
    ind_inc_t_py = D.get('ind_inc_py', None)
    tot_inc_py   = (sales_tot_py + ind_inc_t_py) if (sales_tot_py is not None and ind_inc_t_py is not None) else None

    s3rows.append([P("I.   INCOME","BOLD"),           N(""), A(None),       A(None)])
    s3rows.append([P("Revenue from Operations","IND"), N("14"), A(sales_tot), A(sales_tot_py)])
    s3rows.append([P("Other Income","IND"),            N("15"), A(ind_inc_t), A(ind_inc_t_py)])
    s3rows.append([P("Total Income  (I)","BOLD"),      N(""),   AB(tot_inc),  AB(tot_inc_py)])
    s3rows.append([P("","NRM"), N(""), A(None), A(None)])

    s3rows.append([P("II.  EXPENSES","BOLD"),          N(""), A(None), A(None)])
    cost_mat = abs(D['purch_rm']) + abs(D['purch_clg'])
    inv_chg  = D['opening_stock_abs'] - D['closing_stock_abs']
    # PY values
    purch_py  = D.get('purch_py', None)
    os_py     = D.get('opening_stock_py', None)
    cs_py     = D.get('closing_stock_py', None)
    inv_chg_py = ((os_py - cs_py) if (os_py is not None and cs_py is not None) else None)
    np_py     = D.get('net_profit_py', None)
    gp_py     = D.get('gross_profit_py', None)
    ind_exp_py = D.get('ind_exp_py', None)
    tot_exp_py = ((gp_py + ind_exp_py - ind_inc_t_py)
                  if (gp_py is not None and ind_exp_py is not None and ind_inc_t_py is not None)
                  else None)
    ebt_py    = (tot_inc_py - tot_exp_py) if (tot_inc_py is not None and tot_exp_py is not None) else None
    pnl_bf_py = D.get('pnl_bf_py', None)

    s3rows.append([P("Cost of Materials Consumed","IND"),          N("16"), A(cost_mat),  A(purch_py)])
    s3rows.append([P("Changes in Inventories  (FG / WIP / Stock)","IND"), N("17"), A(inv_chg), A(inv_chg_py)])
    if D['purch_job']  > 0: s3rows.append([P("Job Work Charges","IND"),   N(""), A(D['purch_job']),  A(None)])
    if D['mfg_exp']    > 0: s3rows.append([P("Manufacturing & Other Direct Expenses","IND"), N("18"), A(D['mfg_exp']), A(None)])
    if D['power_fuel'] > 0: s3rows.append([P("Power and Fuel","IND"),     N(""), A(D['power_fuel']), A(None)])
    sc_tot = D['stores_exp'] + D['stores_cons']
    if sc_tot          > 0: s3rows.append([P("Stores & Consumables Consumed","IND"), N(""), A(sc_tot), A(None)])
    if D['emp_exp']    > 0: s3rows.append([P("Employee Benefit Expense","IND"),      N("19"), A(D['emp_exp']),    A(None)])
    if D['finance_exp']> 0: s3rows.append([P("Finance Costs","IND"),                N("20"), A(D['finance_exp']),A(None)])
    other_e = D['admin_exp']+D['factory_exp']+D['legal_exp']+D['selling_exp']
    if other_e         > 0: s3rows.append([P("Other Expenses","IND"),               N("21"), A(other_e),         A(None)])

    tot_exp = (cost_mat + inv_chg + D['purch_job'] + D['mfg_exp'] + D['power_fuel'] +
               sc_tot + D['emp_exp'] + D['finance_exp'] + other_e)
    s3rows.append([P("Total Expenses  (II)","BOLD"),  N(""), AB(tot_exp), AB(tot_exp_py)])
    s3rows.append([P("","NRM"), N(""), A(None), A(None)])

    ebt = tot_inc - tot_exp
    s3rows.append([P("III.  Profit / (Loss) before Tax  (I − II)","BOLD"), N(""), AB(ebt), AB(ebt_py)])
    tax = D['income_tax']
    s3rows.append([P("IV.   Tax Expense","IND"),            N(""), A(tax),         A(None)])
    nat = ebt - tax
    nat_py = (ebt_py - tax) if ebt_py is not None else None
    s3rows.append([P("V.    Net Profit / (Loss) after Tax","BOLD"), N(""), AB(nat), AB(nat_py)])
    s3rows.append([P("      Add: P&L balance brought forward","IND"), N(""), A(D['pnl_bf']), A(pnl_bf_py)])
    s3rows.append([P("VI.   Balance carried to Balance Sheet","BOLD"), N(""),
                   AB(nat + D['pnl_bf']),
                   AB((nat_py + pnl_bf_py) if (nat_py is not None and pnl_bf_py is not None) else None)])

    sc = D['equity_share']
    if sc > 0:
        eps = nat / (sc / 10)
        eps_py = (nat_py / (sc / 10)) if nat_py is not None else None
        s3rows.append([P("","NRM"), N(""), A(None), A(None)])
        s3rows.append([P("Earnings Per Share — Basic & Diluted (Face Value ₹10)","IND"),
                        N(""),
                        Paragraph(f"₹ {eps:,.2f}", STYLES["AMT"]),
                        Paragraph(f"₹ {eps_py:,.2f}" if eps_py is not None else "-", STYLES["AMT"])])

    ts4 = [
        ("FONTSIZE",   (0,0),(-1,-1),7.5),
        ("TOPPADDING", (0,0),(-1,-1),2.5),("BOTTOMPADDING",(0,0),(-1,-1),2.5),
        ("LEFTPADDING",(0,0),(0,-1),4),("RIGHTPADDING",(-1,0),(-1,-1),4),
        ("VALIGN",     (0,0),(-1,-1),"MIDDLE"),
        ("GRID",       (0,0),(-1,-1),0.2,C_BORD),
    ]
    for i, row in enumerate(s3rows):
        txt = row[0].text if hasattr(row[0],'text') else ''
        if 'Net Profit' in txt and 'after Tax' in txt:
            ts4 += [("BACKGROUND",(0,i),(-1,i),C_DARK),("TEXTCOLOR",(0,i),(-1,i),colors.white),
                    ("FONTNAME",(0,i),(-1,i),"Helvetica-Bold")]
        elif txt.startswith('Total') or 'Balance carried' in txt:
            ts4 += [("BACKGROUND",(0,i),(-1,i),C_MID),("TEXTCOLOR",(0,i),(-1,i),colors.white),
                    ("FONTNAME",(0,i),(-1,i),"Helvetica-Bold")]
        elif txt.startswith('I.') or txt.startswith('II.'):
            ts4 += [("BACKGROUND",(0,i),(-1,i),colors.HexColor("#F0F0F0")),
                    ("FONTNAME",(0,i),(-1,i),"Helvetica-Bold")]
        elif i % 2 == 0:
            ts4 += [("BACKGROUND",(0,i),(-1,i),C_LIGHT)]
    story.append(Table(s3rows, colWidths=CW4, style=TableStyle(ts4)))

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — NOTES TO ACCOUNTS  (all dynamic from TB data)
# ══════════════════════════════════════════════════════════════════════════════
def build_notes(D, story):
    CW3 = [PAGE_W*0.60, PAGE_W*0.20, PAGE_W*0.20]
    fy_e= D['fy_end']

    story.append(PageBreak())
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(D['company'], STYLES["H1"]))
    story.append(Paragraph("NOTES TO THE FINANCIAL STATEMENTS",
        ParagraphStyle("T",parent=STYLES["H1"],fontSize=10,spaceAfter=1)))
    story.append(Paragraph(f"As at / For the year ended  {fy_e}", STYLES["H2"]))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_DARK, spaceAfter=5))

    def hdr(rows):
        return [[P(f"<b>{r}</b>","BOLD") for r in rows]]

    # NOTE 1 — Share Capital
    note_hdr(1, "Share Capital", story)
    face = 10
    shares = int(abs(D['equity_share']) / face)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Authorised Capital",1), A(None),                    A(None)],
        [P("Issued, Subscribed & Paid-up",1),
         Paragraph(f"{shares:,} Equity Shares of ₹{face} each", STYLES["AMT"]),
         A(None)],
        [P("Amount",1),             A(abs(D['equity_share'])),  A(None)],
        [P("<b>Total</b>","BOLD"),  AB(abs(D['equity_share'])), AB(None)],
    ], CW3, story)

    # NOTE 2 — Reserves & Surplus
    note_hdr(2, "Reserves & Surplus", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("General Reserve / Retained Earnings",1), A(abs(D['reserves'])),  A(None)],
        [P("Add: Profit / (Loss) for the year",1),   A(D['net_profit']),     A(None)],
        [P("Add: P&L A/c balance b/f",1),            A(D['pnl_bf']),         A(None)],
        [P("<b>Total</b>","BOLD"),
         AB(abs(D['reserves'])+D['net_profit']+D['pnl_bf']),                  AB(None)],
    ], CW3, story)

    # NOTE 3 — Long-term Borrowings
    note_hdr(3, "Long-term Borrowings (Secured)", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Term Loans from Banks / NBFCs",1), A(abs(D['secured_loans'])), A(None)],
        [P("<b>Total</b>","BOLD"),             AB(abs(D['secured_loans'])), AB(None)],
    ], CW3, story)

    # NOTE 4 — Short-term Borrowings
    note_hdr(4, "Short-term Borrowings (Bank OD / CC)", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Cash Credit / OD (Secured against stock & book debts)",1),
         A(abs(D['bank_od'])), A(None)],
        [P("<b>Total</b>","BOLD"), AB(abs(D['bank_od'])), AB(None)],
    ], CW3, story)

    # NOTE 5 — Trade Payables
    note_hdr(5, "Trade Payables (Sundry Creditors)", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Creditors (Micro & Small Enterprises — refer MSME disclosure)",1), A(None), A(None)],
        [P("Other Creditors",1), A(abs(D['sundry_cred'])), A(None)],
        [P("<b>Total</b>","BOLD"), AB(abs(D['sundry_cred'])), AB(None)],
    ], CW3, story)

    # NOTE 6 — Other Current Liabilities
    note_hdr(6, "Other Current Liabilities", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Outstanding Expenses",1),   A(abs(D['outstanding_e'])), A(None)],
        [P("Statutory Liabilities",1),  A(abs(D['stat_liab'])),     A(None)],
        [P("<b>Total</b>","BOLD"),
         AB(abs(D['outstanding_e'])+abs(D['stat_liab'])),           AB(None)],
    ], CW3, story)

    # NOTE 7 — Short-term Provisions
    note_hdr(7, "Short-term Provisions", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Provision for Employee Benefits / Remuneration",1),
         A(abs(D['provisions'])), A(None)],
        [P("<b>Total</b>","BOLD"), AB(abs(D['provisions'])), AB(None)],
    ], CW3, story)

    # NOTE 8 — Fixed Assets
    note_hdr(8, "Fixed Assets  (Net Block)", story)
    fa_gross = abs(D['fa_gross'])
    dep      = D['dep_reserve']
    fa_net   = fa_gross - dep
    # Individual asset lines dynamically from TB
    fa_lines = []
    TB = D['TB']
    for name, val in TB.items():
        if name in ('Fixed Assets', 'DEPRICIATION RESERVE'): continue
        if D['PARENT'].get(name) == 'Fixed Assets' and val != 0:
            fa_lines.append((name, abs(val)))
    fa_lines.sort(key=lambda x: x[1], reverse=True)
    rows8 = hdr(["Asset Description","Gross Block","Net Block"])
    for n, v in fa_lines[:25]:
        rows8.append([P(n,1), A(v), A(None)])
    rows8 += [
        [P("Less: Accumulated Depreciation Reserve",1), A(dep), A(dep)],
        [P("<b>Net Fixed Assets</b>","BOLD"), AB(fa_gross), AB(fa_net)],
    ]
    data_tbl(rows8, CW3, story, total_last=False)

    # NOTE 9 — Inventories (from Stock-in-Hand ledgers — matches Tally exactly)
    note_hdr(9, "Inventories  (Closing Stock)", story)
    rows9 = hdr(["Stock Ledger","Closing Value","Opening Value"])
    for sl in sorted(D.get('stk_ledgers', []), key=lambda x: x['closing'], reverse=True):
        if sl['closing'] > 0 or sl['opening'] > 0:
            rows9.append([P(sl['name'],1), A(sl['closing']), A(sl['opening'])])
    rows9.append([P("<b>Total Closing Stock</b>","BOLD"),
                  AB(D['closing_stock_abs']), AB(D['opening_stock_abs'])])
    data_tbl(rows9, CW3, story, total_last=False)

    # NOTE 10 — Debtors
    note_hdr(10, "Trade Receivables  (Sundry Debtors)", story)
    rows10 = hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Outstanding for more than 6 months (unsecured, considered good)",1), A(None),                    A(None)],
        [P("Other Debts (unsecured, considered good)",1),                        A(abs(D['debtors_net'])),   A(None)],
        [P("<b>Total</b>","BOLD"),                                                AB(abs(D['debtors_net'])),  AB(None)],
    ]
    data_tbl(rows10, CW3, story)
    if D['debtors_list']:
        story.append(Spacer(1,1*mm))
        story.append(Paragraph("<b>Top Debtors by outstanding balance:</b>", STYLES["BOLD"]))
        top10 = sorted(D['debtors_list'], key=lambda x: abs(x['balance']), reverse=True)[:10]
        rows_d = hdr(["Party Name","Group","Balance"])
        for dr in top10:
            rows_d.append([P(dr['name']), P(dr['parent']), A(abs(dr['balance']))])
        data_tbl(rows_d, CW3, story, total_last=False)

    # NOTE 11 — Cash & Bank
    note_hdr(11, "Cash and Cash Equivalents", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Cash in Hand",1),                          A(abs(D['cash_hand'])),   A(None)],
        [P("Balances with Banks — Current Accounts",1),A(abs(D['bank_accts'])),  A(None)],
        [P("<b>Total</b>","BOLD"),
         AB(abs(D['cash_hand'])+abs(D['bank_accts'])),                           AB(None)],
    ], CW3, story)

    # NOTE 12 — Loans & Advances
    note_hdr(12, "Short-term Loans & Advances", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Staff Loans / Advances to Others",1), A(abs(D['loans_adv'])), A(None)],
        [P("<b>Total</b>","BOLD"),                AB(abs(D['loans_adv'])), AB(None)],
    ], CW3, story)

    # NOTE 13 — Other Current Assets
    note_hdr(13, "Other Current Assets", story)
    oca_items = [
        ("Fixed Deposits & Security Deposits",              D['deposits']),
        ("Advance Income Tax / TDS Receivable",             D['taxation_a']),
        ("DBK Receivable (Export)",                         D['dbk']),
        ("RODTEP Incentive Receivable",                     D['rodtep']),
        ("RODTEP Licence in Hand",                          D['rodtep_lic']),
        ("Differential Duty Recoverable (Customs)",         D['diff_duty']),
        ("EMD and Security Deposits",                       D['emda']),
        ("Prepaid Expenses & Other Current Assets",         D['other_ca']),
    ]
    oca_rows = hdr(["Particulars","Current Year","Previous Year"])
    oca_total = 0
    for lbl, v in oca_items:
        if abs(v) > 0:
            oca_rows.append([P(lbl,1), A(abs(v)), A(None)])
            oca_total += abs(v)
    oca_rows.append([P("<b>Total</b>","BOLD"), AB(oca_total), AB(None)])
    data_tbl(oca_rows, CW3, story)

    # NOTE 14 — Revenue
    note_hdr(14, "Revenue from Operations", story)
    sal_rows = hdr(["Particulars","Current Year","Previous Year"])
    if abs(D['sales_domestic']): sal_rows.append([P("Domestic Sales",1),  A(abs(D['sales_domestic'])), A(None)])
    if abs(D['sales_export']):   sal_rows.append([P("Export Sales",1),    A(abs(D['sales_export'])),   A(None)])
    if abs(D['sales_jobwork']):  sal_rows.append([P("Job Work Sales",1),  A(abs(D['sales_jobwork'])),  A(None)])
    # catch any other sales sub-groups dynamically
    accounted = {D['sales_domestic'], D['sales_export'], D['sales_jobwork']}
    for name, val in D['TB'].items():
        if abs(val) > 0 and val not in accounted and D['get_primary'](name) == 'Sales Accounts' and name != 'Sales Accounts':
            if D['PARENT'].get(name) == 'Sales Accounts':
                sal_rows.append([P(name,1), A(abs(val)), A(None)])
    sal_rows.append([P("<b>Total</b>","BOLD"), AB(abs(D['sales'])), AB(None)])
    data_tbl(sal_rows, CW3, story)

    # NOTE 15 — Other Income (dynamic)
    note_hdr(15, "Other Income", story)
    inc_rows = hdr(["Particulars","Current Year","Previous Year"])
    for name, val in sorted(D['ind_inc_lines'], key=lambda x: abs(x[1]), reverse=True):
        inc_rows.append([P(name,1), A(abs(val)), A(None)])
    if not D['ind_inc_lines']:
        inc_rows.append([P("Other Income",1), A(abs(D['ind_inc'])), A(None)])
    inc_rows.append([P("<b>Total</b>","BOLD"), AB(abs(D['ind_inc'])), AB(None)])
    data_tbl(inc_rows, CW3, story)

    # NOTE 16 — Cost of Materials
    note_hdr(16, "Cost of Materials Consumed", story)
    mat_rows = hdr(["Particulars","Current Year","Previous Year"])
    for name, val in D['purch_lines']:
        mat_rows.append([P(name,1), A(abs(val)), A(None)])
    if not D['purch_lines']:
        mat_rows.append([P("Raw Material Purchases",1), A(abs(D['purch_rm'])), A(None)])
    mat_rows.append([P("<b>Total</b>","BOLD"),
                     AB(abs(D['purch_rm'])+abs(D['purch_clg'])), AB(None)])
    data_tbl(mat_rows, CW3, story)

    # NOTE 17 — Changes in Inventories
    note_hdr(17, "Changes in Inventories", story)
    inv_chg = D['opening_stock_abs'] - D['closing_stock_abs']
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Opening Stock",1),         A(D['opening_stock_abs']),  A(None)],
        [P("Less: Closing Stock",1),   A(D['closing_stock_abs']),  A(None)],
        [P("<b>(Increase) / Decrease in Stock</b>","BOLD"), AB(inv_chg), AB(None)],
    ], CW3, story, total_last=False)

    # NOTE 18 — Manufacturing Expenses (dynamic)
    note_hdr(18, "Manufacturing & Other Direct Expenses", story)
    mfg_rows = hdr(["Particulars","Current Year","Previous Year"])
    for name, val in sorted(D['dir_exp_lines'], key=lambda x: abs(x[1]), reverse=True):
        if name not in ('Opening Stock',):
            mfg_rows.append([P(name,1), A(abs(val)), A(None)])
    if not D['dir_exp_lines']:
        mfg_rows.append([P("Manufacturing & Direct Expenses",1), A(abs(D['mfg_exp'])), A(None)])
    mfg_rows.append([P("<b>Total</b>","BOLD"), AB(abs(D['dir_exp'])), AB(None)])
    data_tbl(mfg_rows, CW3, story)

    # NOTE 19 — Employee Expenses
    note_hdr(19, "Employee Benefit Expense", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Salaries, Wages, Bonus & Allowances",1), A(abs(D['emp_exp'])), A(None)],
        [P("<b>Total</b>","BOLD"),                   AB(abs(D['emp_exp'])), AB(None)],
    ], CW3, story)

    # NOTE 20 — Finance Costs
    note_hdr(20, "Finance Costs", story)
    data_tbl(hdr(["Particulars","Current Year","Previous Year"]) + [
        [P("Interest on CC / Term Loans / LC / BG",1), A(abs(D['finance_exp'])), A(None)],
        [P("<b>Total</b>","BOLD"),                      AB(abs(D['finance_exp'])), AB(None)],
    ], CW3, story)

    # NOTE 21 — Other Expenses (dynamic — built from ind_exp_lines)
    # Includes only SUB-GROUPS of Indirect Expenses (e.g. Administrative Exps,
    # Factory Expenses, Legal, Selling) — not direct ledgers booked under the group
    # (e.g. Depreciation, FX, Custom Duty, Income Tax — those are P&L line items).
    # Excludes sub-groups already in Note 19 (Employees) and Note 20 (Finance).
    note_hdr(21, "Other Expenses", story)
    exp_rows = hdr(["Particulars","Current Year","Previous Year"])

    _excl_patterns = [
        'payment to employ', 'salary', 'salaries', 'wages',     # Note 19
        'financial expense', 'finance cost', 'finance charges',  # Note 20
    ]
    def _is_excluded(name):
        return any(p in name.lower() for p in _excl_patterns)

    # PARENT dict from parse_data — tells us which names are sub-groups
    _PARENT = D.get('PARENT', {})

    ot = 0
    for name, val in sorted(D['ind_exp_lines'], key=lambda x: abs(x[1]), reverse=True):
        if _is_excluded(name):
            continue
        # Only show sub-groups (entries that exist in the group hierarchy)
        # Direct ledgers booked under Indirect Expenses are excluded from Note 21
        if name not in _PARENT:
            continue
        display = abs(val)
        if display < 0.01:
            continue
        exp_rows.append([P(name, 1), A(display), A(None)])
        ot += display

    exp_rows.append([P("<b>Total</b>","BOLD"), AB(ot), AB(None)])
    data_tbl(exp_rows, CW3, story)
    data_tbl(exp_rows, CW3, story)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10 — ACCOUNTING POLICIES PAGE
# ══════════════════════════════════════════════════════════════════════════════
def build_policies(D, story):
    story.append(PageBreak())
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(D['company'], STYLES["H1"]))
    story.append(Paragraph("SIGNIFICANT ACCOUNTING POLICIES",
        ParagraphStyle("T",parent=STYLES["H1"],fontSize=10,spaceAfter=1)))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_DARK, spaceAfter=6))

    policies = [
        ("1. Basis of Preparation",
         "The financial statements have been prepared in accordance with the Generally Accepted "
         "Accounting Principles in India (Indian GAAP) and comply with the Accounting Standards "
         "specified under Section 133 of the Companies Act, 2013 read with Rule 7 of the Companies "
         "(Accounts) Rules, 2014 and the relevant provisions of the Companies Act, 2013."),
        ("2. Revenue Recognition",
         "Revenue from sale of goods is recognised when the significant risks and rewards of "
         "ownership are transferred to the buyer. Export sales are recognised on the date of "
         "shipment bill of lading. Interest income is recognised on time-proportion basis."),
        ("3. Inventories",
         "Inventories are valued at lower of cost or net realisable value. Cost of raw materials "
         "is determined on weighted average basis. Finished goods include material cost plus "
         "appropriate share of overheads. Stock values are sourced directly from TallyPrime "
         "inventory register as at the balance sheet date."),
        ("4. Fixed Assets & Depreciation",
         "Fixed assets are stated at historical cost less accumulated depreciation. Depreciation "
         "is provided on the Written Down Value (WDV) method at the rates and in the manner "
         "prescribed under Schedule II of the Companies Act, 2013."),
        ("5. Foreign Currency Transactions",
         "Foreign currency transactions are recorded at the exchange rate prevailing on the date "
         "of transaction. Monetary assets and liabilities denominated in foreign currency are "
         "restated at the rate prevailing at the year-end. Exchange differences are recognised in "
         "the Statement of Profit and Loss."),
        ("6. Taxation",
         "Current tax is provided on the basis of taxable income computed in accordance with "
         "the Income Tax Act, 1961. Deferred tax is recognised on timing differences between "
         "accounting income and taxable income using the liability method at enacted tax rates."),
        ("7. Borrowing Costs",
         "Borrowing costs attributable to acquisition or construction of qualifying assets are "
         "capitalised as part of the cost of such assets up to the date when they are ready for "
         "intended use. All other borrowing costs are charged to the Statement of Profit and Loss."),
        ("8. Export Incentives",
         "Export incentives (Duty Drawback, RODTEP etc.) are recognised in the year in which "
         "the export is made, on accrual basis, at the rate notified by the Government."),
    ]
    for title, text in policies:
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph(f"<b>{title}</b>", STYLES["BOLD"]))
        story.append(Paragraph(text, STYLES["NRM"]))
        story.append(Spacer(1, 1*mm))

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 11 — TALLY REQUIREMENTS PAGE
# ══════════════════════════════════════════════════════════════════════════════
def build_tally_requirements(story):
    story.append(PageBreak())
    story.append(Spacer(1,3*mm))
    story.append(Paragraph("TALLY CONFIGURATION REQUIREMENTS", STYLES["H1"]))
    story.append(Paragraph(
        "What the client must enable in TallyPrime for this report to work correctly",
        STYLES["H2"]))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_DARK, spaceAfter=5))

    sections = [
        ("MANDATORY — Without these the report will fail or show wrong data", colors.HexColor("#C00000"), [
            ("F11 → Accounting Features",  "Maintain Accounts",                     "Yes",  "Core ledger data — without this, no Trial Balance"),
            ("F11 → Accounting Features",  "Enable Bill-wise Entry",                "Yes",  "Required for Debtors / Creditors outstanding bills"),
            ("F11 → Accounting Features",  "Enable Interest Calculation",           "Yes",  "Needed if loan interest is tracked"),
            ("Gateway → F12 → Configure",  "Enable Browser Access for Reports",     "Yes",  "Enables HTTP XML server on port 9000 — this is what the script connects to"),
            ("Gateway → F12 → Configure",  "TDL Config → HTTP Port",               "9000", "Must match --port argument (default 9000)"),
        ]),
        ("REQUIRED FOR INVENTORY (Stock) DATA", colors.HexColor("#C47A00"), [
            ("F11 → Inventory Features",   "Maintain Inventory",                    "Yes",  "Without this, Closing Stock = 0 in the report"),
            ("F11 → Inventory Features",   "Integrate Accounts with Inventory",     "Yes",  "Links stock values to Balance Sheet automatically"),
        ]),
        ("REQUIRED FOR GST DATA", colors.HexColor("#1F5C1F"), [
            ("F11 → Taxation",             "Enable Goods & Services Tax (GST)",     "Yes",  "Required for GST ledger classification"),
            ("F11 → Taxation",             "Set/Alter Company GST Details",         "Fill", "Company GSTIN must be set for report header"),
        ]),
        ("RECOMMENDED — Improves report quality", colors.HexColor("#1F3864"), [
            ("Company Master → F11",       "Company GSTIN",                         "Fill", "Shows in report header"),
            ("Company Master → F11",       "PAN Number",                            "Fill", "Needed for tax compliance fields"),
            ("Company Master → F11",       "CIN Number",                            "Fill", "Required for Companies Act compliance header"),
            ("F11 → Payroll Features",     "Maintain Payroll",                      "Yes",  "Enables employee expense details in Note 19"),
            ("F11 → Accounting Features",  "Enable Cost Centres",                   "Yes",  "Enables department-wise expense breakup"),
        ]),
    ]

    CW_R = [PAGE_W*0.18, PAGE_W*0.36, PAGE_W*0.10, PAGE_W*0.36]

    for sec_title, sec_color, items in sections:
        story.append(Spacer(1, 3*mm))
        story.append(Table([[Paragraph(f"<b>{sec_title}</b>",
            ParagraphStyle("SH", parent=STYLES["BOLD"], textColor=sec_color))]],
            colWidths=[PAGE_W],
            style=TableStyle([
                ("BACKGROUND",(0,0),(-1,-1), colors.HexColor("#F5F5F5")),
                ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
                ("LEFTPADDING",(0,0),(-1,-1),6),
                ("LINEBELOW",(0,0),(-1,0),1,sec_color),
            ])))

        rows = [[
            Paragraph("<b>Where to find it</b>",  STYLES["BOLD"]),
            Paragraph("<b>Setting Name</b>",       STYLES["BOLD"]),
            Paragraph("<b>Set to</b>",             STYLES["BOLD"]),
            Paragraph("<b>Why it matters</b>",     STYLES["BOLD"]),
        ]]
        for where, setting, value, why in items:
            rows.append([
                Paragraph(where,   STYLES["SML"]),
                Paragraph(setting, STYLES["NRM"]),
                Paragraph(f"<b>{value}</b>",
                    ParagraphStyle("V",parent=STYLES["BOLD"],
                                   textColor=colors.HexColor("#1A7A4A") if value=="Yes" else C_DARK)),
                Paragraph(why,     STYLES["SML"]),
            ])

        ts = [
            ("FONTSIZE",(0,0),(-1,-1),7.5),
            ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("LEFTPADDING",(0,0),(0,-1),5),
            ("GRID",(0,0),(-1,-1),0.2,C_BORD),
            ("BACKGROUND",(0,0),(-1,0),C_DARK),
            ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ]
        for i in range(1,len(rows),2):
            ts.append(("BACKGROUND",(0,i),(-1,i),C_LIGHT))
        story.append(Table(rows, colWidths=CW_R, style=TableStyle(ts)))

    # Quick-start note
    story.append(Spacer(1,4*mm))
    story.append(HRFlowable(width=PAGE_W, thickness=0.5, color=C_BORD))
    story.append(Spacer(1,2*mm))
    story.append(Paragraph("<b>Quick Verification in Tally:</b>", STYLES["BOLD"]))
    steps = [
        "Open TallyPrime → Load the company you want to report on",
        "Press F11 → Check all MANDATORY settings above are set to Yes",
        "Go to Gateway of Tally → press F12 → Enable Browser Access for Reports = Yes",
        "Note the HTTP port shown (default 9000)",
        "Open a browser and go to  http://localhost:9000  — if Tally shows a response, connection is working",
        "Now run:  python tally_sch3.py --list-companies  to confirm your company is visible",
        "Then run:  python tally_sch3.py --company \"YOUR COMPANY\" --year 2026  to generate the report",
    ]
    for i, s in enumerate(steps, 1):
        story.append(Paragraph(f"  {i}.  {s}", STYLES["NRM"]))

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 12 — MASTER PDF GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
def generate_pdf(D, out_path):
    global STYLES
    STYLES = mk_styles()

    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            topMargin=12*mm, bottomMargin=12*mm,
                            leftMargin=10*mm, rightMargin=10*mm)
    story = []
    build_tb(D, story)          # Page 1+  : Trial Balance
    story.append(PageBreak())
    build_bs(D, story)          # Page next : Balance Sheet
    build_pnl(D, story)         # Page next : P&L Account
    build_notes(D, story)       # Page next : Notes 1-21
    build_policies(D, story)    # Page next : Accounting Policies
    build_tally_requirements(story)  # Last page : Tally Setup Guide
    doc.build(story)
    print(f"  PDF saved  → {out_path}")
    return out_path

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 13 — VOUCHER SCRAPER  (direct from Tally — no DB required)
# Fetches vouchers via Tally XML API using 1-day chunks + dual-report strategy
# (Voucher Register + Day Book) to avoid truncation and catch all voucher types.
# ══════════════════════════════════════════════════════════════════════════════

def _parse_amount(text):
    """Parse Tally amount string → signed float. Tally sends negatives as '-123.45'."""
    if not text:
        return 0.0
    try:
        return float(str(text).replace(",", "").strip())
    except Exception:
        return 0.0

def _parse_date_str(raw):
    """Parse YYYYMMDD / DD-MM-YYYY / YYYY-MM-DD / DD/MM/YYYY → date or None."""
    raw = (raw or "").strip()
    for fmt in ("%Y%m%d", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except Exception:
            pass
    return None

def _parse_voucher_xml(raw_bytes):
    """
    Parse Tally XML (Voucher Register / Day Book / Stock Journal response).

    Returns list of voucher dicts. Each dict matches exactly what Tally shows
    in its Day Book / Ledger Voucher view:

    Header fields
    ─────────────
      number        : voucher number (auto-generated if blank)
      date          : YYYY-MM-DD
      type          : voucher type name (Sales, Purchase, Payment …)
      narration     : voucher narration
      is_cancelled  : bool
      is_optional   : bool (memo / non-posting)
      guid          : Tally MASTERID — used for dedup
      ref_number    : manual reference number (ALTEREDON or REFERNCENO)
      place_of_supply: state name for GST

    Ledger entries (accounting side) — entry_type DR/CR matches Tally Day Book
    ─────────────────────────────────────────────────────────────────────────────
      entries: [{
        ledger      : ledger name
        amount      : absolute float
        type        : 'DR' or 'CR'
        gstin       : party GSTIN if available (for GST reconciliation)
        bill_refs   : [{ name, amount, type:'New'|'Agst Ref'|'Advance', due_date }]
        cost_centres: [{ name, amount }]
      }]

    Inventory items (stock side) — for Sales, Purchase, Stock Journal etc.
    ───────────────────────────────────────────────────────────────────────
      items: [{
        name        : stock item name
        qty         : quantity with unit (e.g. '10 Nos')
        rate        : rate per unit
        amount      : line value
        godown      : godown / warehouse name
        batch       : batch name if any
        hsn         : HSN/SAC code
      }]

    Cheque / bank details
    ──────────────────────
      cheque_number : cheque no. (from CHEQUENUMBER or narration)
      cheque_date   : cheque date
      bank_name     : bank ledger name (largest CR on Payment)
    """
    try:
        root = ET.fromstring(fix_enc(raw_bytes))
    except Exception:
        return []

    vouchers = []
    for v in root.iter("VOUCHER"):
        # ── Header ──────────────────────────────────────────────────────────
        vnum   = (v.findtext("VOUCHERNUMBER") or "").strip()
        vraw   = (v.findtext("DATE") or v.get("DATE") or "").strip()
        vtype  = (v.findtext("VOUCHERTYPENAME") or v.get("VOUCHERTYPENAME") or "").strip()
        narr   = (v.findtext("NARRATION") or "").strip()
        ref_no = (v.findtext("REFERENCE") or v.findtext("REFERNCENO") or
                  v.findtext("ALTEREDON") or "").strip()
        pos    = (v.findtext("PLACEOFSUPPLY") or v.findtext("STATENAME") or "").strip()

        guid   = (v.get("MASTERID") or v.get("GUID") or
                  v.findtext("MASTERID") or v.findtext("GUID") or
                  v.findtext("ALTERID") or "").strip()

        is_cancelled = (v.findtext("ISCANCELLED") or v.get("ISCANCELLED") or "").strip().upper() == "YES"
        is_optional  = (v.findtext("ISOPTIONAL")  or v.get("ISOPTIONAL")  or "").strip().upper() == "YES"

        # Cheque details
        cheque_number = (v.findtext("CHEQUENUMBER") or v.findtext("CHEQUENO") or "").strip()
        cheque_date   = _parse_date_str(v.findtext("CHEQUEDATE") or "")

        vdate = _parse_date_str(vraw)
        if vdate is None:
            continue

        # ── Ledger entries ───────────────────────────────────────────────────
        entries = []
        for tag in ("ALLLEDGERENTRIES.LIST", "LEDGERENTRIES.LIST",
                    "ALLLEDGERENTRIES",       "LEDGERENTRIES"):
            for entry in v.findall(tag):
                ledger = (entry.findtext("LEDGERNAME") or "").strip()
                if not ledger:
                    continue

                raw_amt = entry.findtext("AMOUNT") or "0"
                amount  = abs(_parse_amount(raw_amt))

                deemed  = (entry.findtext("ISDEEMEDPOSITIVE") or "").strip().lower()
                etype   = "DR" if deemed in ("yes", "true", "1") else "CR"

                # Party GSTIN (present on party ledger entry in GST vouchers)
                gstin = (entry.findtext("PARTYLEDGERGSTIN") or
                         entry.findtext("GSTREGISTRATIONTYPE") or
                         entry.findtext("PARTYGSTIN") or "").strip()

                # Bill-wise references (outstanding / ageing)
                bill_refs = []
                for ba in entry.findall("BILLALLOCATIONS.LIST") or []:
                    bname  = (ba.findtext("NAME") or "").strip()
                    bamt   = _parse_amount(ba.findtext("AMOUNT") or "0")
                    btype  = (ba.findtext("BILLTYPE") or "New Ref").strip()
                    bdue   = _parse_date_str(ba.findtext("DUEDATE") or "")
                    if bname:
                        bill_refs.append({
                            "name":     bname,
                            "amount":   abs(bamt),
                            "type":     btype,
                            "due_date": str(bdue) if bdue else "",
                        })

                # Cost centre allocations
                cost_centres = []
                for ca in entry.findall("CATEGORYALLOCATIONS.LIST") or []:
                    for cc in ca.findall("COSTCENTREALLOCATIONS.LIST") or []:
                        ccname = (cc.findtext("NAME") or "").strip()
                        ccamt  = abs(_parse_amount(cc.findtext("AMOUNT") or "0"))
                        if ccname:
                            cost_centres.append({"name": ccname, "amount": ccamt})

                entries.append({
                    "ledger":       ledger,
                    "amount":       round(amount, 2),
                    "type":         etype,
                    "gstin":        gstin,
                    "bill_refs":    bill_refs,
                    "cost_centres": cost_centres,
                })

        # ── Inventory / stock item lines ─────────────────────────────────────
        items = []
        for tag in ("ALLINVENTORYENTRIES.LIST", "INVENTORYENTRIES.LIST",
                    "ALLINVENTORYENTRIES",       "INVENTORYENTRIES"):
            for ie in v.findall(tag):
                iname   = (ie.findtext("STOCKITEMNAME") or ie.findtext("ITEMNAME") or "").strip()
                if not iname:
                    continue
                # Quantity: BILLEDQTY preferred (as invoiced), else ACTUALQTY
                qty_raw  = (ie.findtext("BILLEDQTY") or ie.findtext("ACTUALQTY") or "").strip()
                rate_raw = (ie.findtext("RATE") or "").strip()
                amt_raw  = (ie.findtext("AMOUNT") or "0").strip()
                godown   = (ie.findtext("GODOWNNAME") or ie.findtext("GODOWN") or "").strip()
                batch    = (ie.findtext("BATCHNAME") or ie.findtext("BATCH") or "").strip()
                hsn      = (ie.findtext("HSNCODE") or ie.findtext("HSNSACCODE") or "").strip()

                items.append({
                    "name":   iname,
                    "qty":    qty_raw,
                    "rate":   rate_raw,
                    "amount": round(abs(_parse_amount(amt_raw)), 2),
                    "godown": godown,
                    "batch":  batch,
                    "hsn":    hsn,
                })

        # ── Derive party name (largest entry by amount for Sales/Purchase) ──
        party = ""
        if entries:
            # For Sales: party = largest DR entry (customer being debited)
            # For Purchase: party = largest CR entry (vendor being credited)
            dr_entries = [e for e in entries if e["type"] == "DR"]
            cr_entries = [e for e in entries if e["type"] == "CR"]
            if vtype in ("Sales", "Credit Note") and dr_entries:
                party = max(dr_entries, key=lambda e: e["amount"])["ledger"]
            elif vtype in ("Purchase", "Debit Note") and cr_entries:
                party = max(cr_entries, key=lambda e: e["amount"])["ledger"]
            elif vtype in ("Receipt",) and dr_entries:
                party = max(dr_entries, key=lambda e: e["amount"])["ledger"]
            elif vtype in ("Payment",) and cr_entries:
                party = max(cr_entries, key=lambda e: e["amount"])["ledger"]

        # Voucher total = sum of DR entries (= sum of CR entries for balanced voucher)
        total_dr = round(sum(e["amount"] for e in entries if e["type"] == "DR"), 2)
        total_cr = round(sum(e["amount"] for e in entries if e["type"] == "CR"), 2)

        vouchers.append({
            "number":          vnum,
            "date":            str(vdate),
            "type":            vtype,
            "narration":       narr,
            "ref_number":      ref_no,
            "place_of_supply": pos,
            "is_cancelled":    is_cancelled,
            "is_optional":     is_optional,
            "guid":            guid,
            "party":           party,
            "cheque_number":   cheque_number,
            "cheque_date":     str(cheque_date) if cheque_date else "",
            "total_dr":        total_dr,
            "total_cr":        total_cr,
            "entries":         entries,
            "items":           items,
        })
    return vouchers


def _dedup_add(v, all_vouchers, seen_guid, seen_fallback, include_cancelled):
    """Add voucher to list if not already seen. Returns True if added."""
    if not include_cancelled and v["is_cancelled"]:
        return False
    if v["is_optional"]:
        return False   # skip memo/non-posting vouchers
    g = v["guid"]
    if g:
        if g in seen_guid:
            return False
        seen_guid.add(g)
    else:
        fk = (v["number"], v["date"], v["type"])
        if fk in seen_fallback:
            return False
        seen_fallback.add(fk)
    all_vouchers.append(v)
    return True


def fetch_vouchers_from_tally(company, from_date, to_date,
                               vtype_filter=None, include_cancelled=True):
    """
    Scrape ALL vouchers from Tally for a date range.

    Strategy (three layers, merged with GUID dedup):
      1. Voucher Register  — accounting vouchers with full ledger entries
      2. Day Book          — all voucher types incl. inventory (Stock Journal, Delivery Note…)
      3. Stock Journal     — direct fetch fallback for inventory-only vouchers

    Each fetch uses 1-day chunks to prevent Tally XML truncation on busy days.

    Returned fields per voucher (matches Tally's own Day Book / Ledger view):
      number, date, type, narration, ref_number, place_of_supply,
      is_cancelled, party, cheque_number, cheque_date, total_dr, total_cr,
      entries: [{ledger, amount, type, gstin, bill_refs, cost_centres}]
      items:   [{name, qty, rate, amount, godown, batch, hsn}]
    """
    from_d = datetime.strptime(from_date, "%Y-%m-%d").date() if isinstance(from_date, str) else from_date
    to_d   = datetime.strptime(to_date,   "%Y-%m-%d").date() if isinstance(to_date,   str) else to_date

    all_vouchers  = []
    seen_guid     = set()
    seen_fallback = set()

    current = from_d
    while current <= to_d:
        ds = tally_date(current)

        # Layer 1 + 2: Voucher Register and Day Book
        for report_name in ("Voucher Register", "Day Book"):
            vtype_clause = (
                f"<VOUCHERTYPENAME>{xe(vtype_filter)}</VOUCHERTYPENAME>"
                if vtype_filter else ""
            )
            xml_body = (
                f"<ENVELOPE>"
                f"<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>"
                f"<BODY><EXPORTDATA><REQUESTDESC>"
                f"<REPORTNAME>{report_name}</REPORTNAME>"
                f"<STATICVARIABLES>"
                f"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
                f"<SVFROMDATE>{ds}</SVFROMDATE><SVTODATE>{ds}</SVTODATE>"
                f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
                f"<EXPLODEFLAG>Yes</EXPLODEFLAG>"
                f"{vtype_clause}"
                f"</STATICVARIABLES>"
                f"</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>"
            )
            try:
                raw = post(xml_body, timeout=60)
                for v in _parse_voucher_xml(raw):
                    _dedup_add(v, all_vouchers, seen_guid, seen_fallback, include_cancelled)
            except Exception as e:
                print(f"  [vouchers] {ds} [{report_name}] skipped: {e}")

        # Layer 3: Stock Journal direct (catches inventory-only vouchers missed above)
        try:
            sj_xml = (
                f"<ENVELOPE>"
                f"<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>"
                f"<BODY><EXPORTDATA><REQUESTDESC>"
                f"<REPORTNAME>Stock Journal</REPORTNAME>"
                f"<STATICVARIABLES>"
                f"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
                f"<SVFROMDATE>{ds}</SVFROMDATE><SVTODATE>{ds}</SVTODATE>"
                f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
                f"<EXPLODEFLAG>Yes</EXPLODEFLAG>"
                f"</STATICVARIABLES>"
                f"</REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>"
            )
            raw_sj = post(sj_xml, timeout=60)
            for v in _parse_voucher_xml(raw_sj):
                _dedup_add(v, all_vouchers, seen_guid, seen_fallback, include_cancelled)
        except Exception:
            pass   # Stock Journal absent in older Tally — skip silently

        current += timedelta(days=1)

    all_vouchers.sort(key=lambda x: (x["date"], x["number"] or ""))
    return all_vouchers


def fetch_ledger_vouchers(company, ledger_name, from_date, to_date):
    """
    Fetch every voucher that touches a specific ledger — exactly like
    Tally's own Ledger Voucher report (Gateway → Display → Account Books → Ledger).

    Uses Tally's Collection API with a LEDGER filter so only vouchers
    that post to `ledger_name` are returned, fully exploded with entries.

    Returns the same voucher dict format as fetch_vouchers_from_tally.
    """
    from_d = datetime.strptime(from_date, "%Y-%m-%d").date() if isinstance(from_date, str) else from_date
    to_d   = datetime.strptime(to_date,   "%Y-%m-%d").date() if isinstance(to_date,   str) else to_date

    xml_body = (
        f"<ENVELOPE><HEADER><VERSION>1</VERSION>"
        f"<TALLYREQUEST>Export</TALLYREQUEST>"
        f"<TYPE>Collection</TYPE><ID>LedVouchers</ID></HEADER>"
        f"<BODY><DESC>"
        f"<STATICVARIABLES>"
        f"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
        f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
        f"<SVFROMDATE>{tally_date(from_d)}</SVFROMDATE>"
        f"<SVTODATE>{tally_date(to_d)}</SVTODATE>"
        f"</STATICVARIABLES>"
        f"<TDL><TDLMESSAGE>"
        f"<COLLECTION NAME=\"LedVouchers\" ISMODIFY=\"No\">"
        f"<TYPE>Voucher</TYPE>"
        f"<BELONGSTO>\"{xe(ledger_name)}\"</BELONGSTO>"
        f"<FETCH>*</FETCH>"
        f"<EXPLODEFLAG>Yes</EXPLODEFLAG>"
        f"</COLLECTION>"
        f"</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        raw = post(xml_body, timeout=120)
        vouchers = _parse_voucher_xml(raw)
        vouchers.sort(key=lambda x: (x["date"], x["number"] or ""))
        return vouchers
    except Exception as e:
        print(f"  [ledger vouchers] {ledger_name}: {e}")
        return []


def fetch_vouchers_by_type(company, voucher_type, from_date, to_date):
    """
    Fetch all vouchers of a specific type for the period using
    Tally's Collection API filtered by VoucherType — more reliable
    than report-based fetch for high-volume types.

    Returns same voucher dict format as fetch_vouchers_from_tally.
    """
    from_d = datetime.strptime(from_date, "%Y-%m-%d").date() if isinstance(from_date, str) else from_date
    to_d   = datetime.strptime(to_date,   "%Y-%m-%d").date() if isinstance(to_date,   str) else to_date

    xml_body = (
        f"<ENVELOPE><HEADER><VERSION>1</VERSION>"
        f"<TALLYREQUEST>Export</TALLYREQUEST>"
        f"<TYPE>Collection</TYPE><ID>TypedVouchers</ID></HEADER>"
        f"<BODY><DESC>"
        f"<STATICVARIABLES>"
        f"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
        f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
        f"<SVFROMDATE>{tally_date(from_d)}</SVFROMDATE>"
        f"<SVTODATE>{tally_date(to_d)}</SVTODATE>"
        f"</STATICVARIABLES>"
        f"<TDL><TDLMESSAGE>"
        f"<COLLECTION NAME=\"TypedVouchers\" ISMODIFY=\"No\">"
        f"<TYPE>Voucher</TYPE>"
        f"<FILTER>FilterByType</FILTER>"
        f"<FETCH>*</FETCH>"
        f"<EXPLODEFLAG>Yes</EXPLODEFLAG>"
        f"</COLLECTION>"
        f"<FUNCTION NAME=\"FilterByType\" PARAMS=\"Voucher\">"
        f"$$VoucherTypeName:Voucher=\"{xe(voucher_type)}\""
        f"</FUNCTION>"
        f"</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        raw = post(xml_body, timeout=120)
        vouchers = _parse_voucher_xml(raw)
        vouchers.sort(key=lambda x: (x["date"], x["number"] or ""))
        return vouchers
    except Exception as e:
        print(f"  [vouchers by type] {voucher_type}: {e}")
        # Fallback to day-chunk fetch for this type
        return fetch_vouchers_from_tally(company, from_date, to_date,
                                          vtype_filter=voucher_type)


def get_voucher_types_from_tally(company):
    """Fetch all voucher type names defined in TallyPrime for this company."""
    xml_body = (
        f"<ENVELOPE><HEADER><VERSION>1</VERSION>"
        f"<TALLYREQUEST>Export</TALLYREQUEST>"
        f"<TYPE>Collection</TYPE><ID>VTypeList</ID></HEADER>"
        f"<BODY><DESC>"
        f"<STATICVARIABLES>"
        f"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
        f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
        f"</STATICVARIABLES>"
        f"<TDL><TDLMESSAGE>"
        f"<COLLECTION NAME=\"VTypeList\" ISMODIFY=\"No\">"
        f"<TYPE>VoucherType</TYPE><FETCH>Name</FETCH>"
        f"</COLLECTION>"
        f"</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    try:
        raw = post(xml_body, timeout=30)
        root = ET.fromstring(fix_enc(raw))
        types = []
        for el in root.iter("VOUCHERTYPE"):
            name = (el.get("NAME") or el.findtext("NAME") or "").strip()
            if name:
                types.append(name)
        return sorted(set(types), key=str.upper)
    except Exception:
        return []


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 14 — FLASK SERVER  (serves index.html + live-scrape API endpoints)
#
# Endpoints served:
#   GET  /                         → index.html
#   GET  /companies                → list companies open in Tally
#   GET  /voucher_types            → voucher types for a company
#   GET  /daybook                  → paginated voucher+entry rows (live Tally)
#   GET  /api/v1/vouchers          → voucher-level list (header only, no entries)
#   GET  /api/v1/summary           → quick counts for dashboard KPIs
#   POST /set_company              → set active company
# ══════════════════════════════════════════════════════════════════════════════

def _make_flask_server(port=5057, tally_port=9000):
    """Build and return the Flask app. Call app.run() to start."""
    try:
        from flask import Flask, request as freq, jsonify, send_file, abort
        from flask_cors import CORS
    except ImportError:
        for pkg in ("flask", "flask-cors"):
            _pip(pkg)
        from flask import Flask, request as freq, jsonify, send_file, abort
        from flask_cors import CORS

    app = Flask(__name__, static_folder=None)
    CORS(app)

    # ── Resolve index.html location (same dir as this script) ────────────────
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    _index_path = os.path.join(_script_dir, "index.html")

    # Update tally URL based on chosen port
    set_tally_url(f"http://localhost:{tally_port}")

    # ── Helper ────────────────────────────────────────────────────────────────
    def _param(key, default=None):
        return freq.args.get(key, default)

    def _require(key):
        v = freq.args.get(key)
        if not v:
            abort(400, description=f"Missing required parameter: {key}")
        return v

    # ── Routes ────────────────────────────────────────────────────────────────

    @app.route("/")
    def index():
        if os.path.exists(_index_path):
            return send_file(_index_path)
        return "<h2>index.html not found — place it next to tally_core.py</h2>", 404

    @app.route("/companies")
    def companies():
        try:
            cos = list_companies()
            return jsonify({"companies": cos})
        except Exception as e:
            return jsonify({"error": str(e), "companies": []}), 500

    @app.route("/voucher_types")
    def voucher_types():
        company = _require("company")
        try:
            types = get_voucher_types_from_tally(company)
            return jsonify({"types": types, "voucher_types": types})
        except Exception as e:
            return jsonify({"error": str(e), "types": []}), 500

    @app.route("/daybook")
    def daybook():
        """
        Live-scrape Tally and return paginated voucher+entry rows.
        Query params:
          company    (required)
          from_date  YYYY-MM-DD  (default: FY start)
          to_date    YYYY-MM-DD  (default: today)
          vtype      voucher type filter (optional)
          page       1-based page number (default 1)
          page_size  rows per page (default 100, max 1000)
        Response: {total, page, total_pages, vouchers:[{date,number,type,narration,is_cancelled,entries:[]}]}
        """
        company   = _require("company")
        today     = date.today()
        fy_start  = str(date(today.year if today.month >= 4 else today.year-1, 4, 1))
        from_date = _param("from_date", fy_start)
        to_date   = _param("to_date",   str(today))
        vtype     = _param("vtype", None)
        try:
            page      = max(1, int(_param("page", 1)))
            page_size = min(1000, max(1, int(_param("page_size", 100))))
        except ValueError:
            page, page_size = 1, 100

        try:
            vouchers = fetch_vouchers_from_tally(
                company, from_date, to_date,
                vtype_filter=vtype,
                include_cancelled=True
            )
        except Exception as e:
            return jsonify({"error": str(e), "vouchers": [], "total": 0}), 500

        # Count at entry level (each entry = 1 table row in Day Book view)
        entry_rows = []
        for v in vouchers:
            if v["entries"]:
                for e in v["entries"]:
                    entry_rows.append({**v, "_entry": e})
            else:
                entry_rows.append({**v, "_entry": None})

        total       = len(entry_rows)
        total_pages = max(1, (total + page_size - 1) // page_size)
        page        = min(page, total_pages)
        sliced      = entry_rows[(page-1)*page_size : page*page_size]

        # Re-group sliced entry rows back into voucher+entries structure
        # (frontend expects {vouchers: [{..., entries:[]}]})
        vmap = {}
        out_vouchers = []
        for row in sliced:
            key = (row["date"], row["number"], row["type"], row["guid"])
            if key not in vmap:
                vobj = {
                    "date":         row["date"],
                    "number":       row["number"],
                    "type":         row["type"],
                    "narration":    row["narration"],
                    "is_cancelled": row["is_cancelled"],
                    "guid":         row["guid"],
                    "entries":      [],
                }
                vmap[key] = vobj
                out_vouchers.append(vobj)
            if row["_entry"]:
                vmap[key]["entries"].append(row["_entry"])

        return jsonify({
            "total":       total,
            "page":        page,
            "page_size":   page_size,
            "total_pages": total_pages,
            "from_date":   from_date,
            "to_date":     to_date,
            "company":     company,
            "vouchers":    out_vouchers,
        })

    @app.route("/api/v1/vouchers")
    def api_vouchers():
        """
        Voucher-level list (header only, no entry drill-down).
        Query params: company, from, to, type, limit (default 500)
        Response: {total, vouchers:[{date,number,type,narration,is_cancelled}]}
        """
        company   = _require("company")
        today     = date.today()
        fy_start  = str(date(today.year if today.month >= 4 else today.year-1, 4, 1))
        from_date = _param("from", fy_start)
        to_date   = _param("to",   str(today))
        vtype     = _param("type", None)
        try:
            limit = min(5000, max(1, int(_param("limit", 500))))
        except ValueError:
            limit = 500

        try:
            vouchers = fetch_vouchers_from_tally(
                company, from_date, to_date,
                vtype_filter=vtype,
                include_cancelled=True
            )
        except Exception as e:
            return jsonify({"error": str(e), "vouchers": [], "total": 0}), 500

        # Return headers only (strip entries for speed; entries via /daybook)
        headers = [
            {k: v for k, v in vch.items() if k != "entries"}
            for vch in vouchers[:limit]
        ]
        return jsonify({
            "total":    len(vouchers),
            "returned": len(headers),
            "from":     from_date,
            "to":       to_date,
            "vouchers": headers,
        })

    @app.route("/api/v1/summary")
    def api_summary():
        """
        Quick stats for a company for the current FY.
        Scrapes last 30 days of vouchers for speed; returns estimated counts.
        """
        company = _require("company")
        today   = date.today()
        fy_start = str(date(today.year if today.month >= 4 else today.year-1, 4, 1))
        # For summary, scrape last 30 days only (fast)
        from_30  = str(today - timedelta(days=30))
        try:
            sample = fetch_vouchers_from_tally(company, from_30, str(today))
            vtypes = list({v["type"] for v in sample if v["type"]})
            return jsonify({
                "company":       company,
                "sample_from":   from_30,
                "sample_to":     str(today),
                "vouchers":      len(sample),
                "entries":       sum(len(v["entries"]) for v in sample),
                "ledgers":       len({e["ledger"] for v in sample for e in v["entries"]}),
                "voucher_types": vtypes,
                "note":          "Counts are for last 30 days. Use /daybook for full range."
            })
        except Exception as e:
            return jsonify({"error": str(e), "vouchers": 0, "entries": 0, "ledgers": 0}), 500

    # ── /status  (health check) ───────────────────────────────────────────────
    @app.route("/status")
    def status():
        try:
            post(b"<ENVELOPE/>", timeout=3)
            tally_ok = True
        except Exception:
            tally_ok = False
        return jsonify({
            "status":    "ok",
            "tally_url": TALLY_URL,
            "tally_ok":  tally_ok,
            "server":    "tally_core.py",
        })

    # ── /sync_status_all  (stub — no DB; returns live Tally connection state) ─
    @app.route("/sync_status_all")
    def sync_status_all():
        try:
            cos = list_companies()
        except Exception:
            cos = []
        return jsonify({
            "synced":    [],           # no DB sync model in this server
            "companies": cos,
            "mode":      "live",       # data is fetched live from Tally, not from DB
        })

    # ── /api/v1/companies  (alias used by some UI paths) ────────────────────
    @app.route("/api/v1/companies")
    def api_companies():
        try:
            cos = list_companies()
            return jsonify({"companies": cos})
        except Exception as e:
            return jsonify({"error": str(e), "companies": []}), 500

    # ── /api/v1/company_period  ───────────────────────────────────────────────
    @app.route("/api/v1/company_period")
    def api_company_period():
        company = _require("company")
        today   = date.today()
        fy_year = today.year if today.month >= 4 else today.year - 1
        return jsonify({
            "company":   company,
            "from_date": str(date(fy_year, 4, 1)),
            "to_date":   str(date(fy_year + 1, 3, 31)),
            "mode":      "live",
        })

    # ── /api/v1/ledgers  — full ledger master with balances live from Tally ──
    @app.route("/api/v1/ledgers")
    def api_ledgers():
        """
        Returns all ledger masters with opening + closing balances.
        Optionally filter by group: ?group=Sundry+Debtors
        Query: company, group (optional), limit (default 1000)
        """
        company  = _require("company")
        grp_filt = _param("group", None)
        try:
            limit = min(5000, max(1, int(_param("limit", 1000))))
        except ValueError:
            limit = 1000

        today    = date.today()
        fy_year  = today.year if today.month >= 4 else today.year - 1
        fy_start = str(date(fy_year, 4, 1))
        fy_end   = str(date(fy_year + 1, 3, 31))

        # Full ledger collection with balances
        body = (
            f"<ENVELOPE><HEADER><VERSION>1</VERSION>"
            f"<TALLYREQUEST>Export</TALLYREQUEST>"
            f"<TYPE>Collection</TYPE><ID>LedgerMasters</ID></HEADER>"
            f"<BODY><DESC>"
            f"<STATICVARIABLES>"
            f"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
            f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
            f"<SVFROMDATE>{tally_date(fy_start)}</SVFROMDATE>"
            f"<SVTODATE>{tally_date(fy_end)}</SVTODATE>"
            f"</STATICVARIABLES>"
            f"<TDL><TDLMESSAGE>"
            f"<COLLECTION NAME=\"LedgerMasters\" ISMODIFY=\"No\">"
            f"<TYPE>Ledger</TYPE>"
            f"<FETCH>Name,Parent,OpeningBalance,ClosingBalance,"
            f"IsBillwiseOn,CreditLimit,GSTRegistrationType,PartyGSTIN,"
            f"PANItNumber,PinCode,CreditPeriod,MasterId</FETCH>"
            f"</COLLECTION>"
            f"</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
        )
        try:
            raw      = post(body, timeout=120)
            root     = ET.fromstring(fix_enc(raw))
            ledgers  = []
            for el in root.iter("LEDGER"):
                name   = (el.get("NAME") or el.findtext("NAME") or "").strip()
                if not name:
                    continue
                parent = (el.findtext("PARENT") or "").strip()
                if grp_filt and parent.lower() != grp_filt.lower():
                    continue
                try:
                    op = float(el.findtext("OPENINGBALANCE") or 0)
                except Exception:
                    op = 0.0
                try:
                    cl = float(el.findtext("CLOSINGBALANCE") or 0)
                except Exception:
                    cl = 0.0
                ledgers.append({
                    "name":          name,
                    "parent":        parent,
                    "opening":       round(op, 2),
                    "closing":       round(cl, 2),
                    "opening_dr_cr": "DR" if op < 0 else ("CR" if op > 0 else ""),
                    "closing_dr_cr": "DR" if cl < 0 else ("CR" if cl > 0 else ""),
                    "gstin":         (el.findtext("PARTYGSTINNO") or el.findtext("PARTYGSTINNUMBER") or "").strip(),
                    "pan":           (el.findtext("PANITNO") or "").strip(),
                    "master_id":     (el.get("MASTERID") or el.findtext("MASTERID") or "").strip(),
                })
            ledgers.sort(key=lambda x: (x["parent"], x["name"]))
            return jsonify({
                "company":  company,
                "fy_start": fy_start,
                "fy_end":   fy_end,
                "total":    len(ledgers),
                "returned": min(len(ledgers), limit),
                "ledgers":  ledgers[:limit],
            })
        except Exception as e:
            return jsonify({"error": str(e), "ledgers": [], "total": 0}), 500

    # ── /api/v1/trial_balance  — live TB from Tally as clean JSON ─────────────
    @app.route("/api/v1/trial_balance")
    def api_trial_balance():
        """
        Returns TB rows exactly as Tally shows them:
        group rows + indented ledger rows, each with DR/CR amounts.
        Query: company, from_date (YYYY-MM-DD), to_date (YYYY-MM-DD)
        """
        company  = _require("company")
        today    = date.today()
        fy_year  = today.year if today.month >= 4 else today.year - 1
        fy_start = _param("from_date", str(date(fy_year, 4, 1)))
        fy_end   = _param("to_date",   str(date(fy_year + 1, 3, 31)))

        try:
            raw_tb = fetch_trial_balance(company, fy_start, fy_end)
        except Exception as e:
            return jsonify({"error": str(e), "rows": []}), 500

        def _find_key(obj, key, depth=0):
            if depth > 8 or not isinstance(obj, dict): return None
            if key in obj: return obj[key]
            for v in obj.values():
                r = _find_key(v, key, depth+1)
                if r is not None: return r
            return None

        def fv(s):
            if s is None: return 0.0
            try: return float(str(s).replace(",", ""))
            except: return 0.0

        tb_names = _find_key(raw_tb, "DSPACCNAME") or []
        tb_info  = _find_key(raw_tb, "DSPACCINFO") or []
        if not isinstance(tb_names, list): tb_names = [tb_names]
        if not isinstance(tb_info,  list): tb_info  = [tb_info]

        # All primary group names — used to set is_group flag
        PRIMARY_GROUPS = {
            "Capital Account","Loans (Liability)","Current Liabilities",
            "Fixed Assets","Current Assets","Sales Accounts",
            "Purchase Accounts","Direct Expenses","Indirect Incomes",
            "Indirect Expenses","Direct Incomes","Investments",
            "Misc. Expenses (ASSET)","Branch / Divisions","Suspense A/c",
            "Profit & Loss A/c","Opening Stock","Stock-in-Hand",
        }

        rows = []
        total_dr = total_cr = 0.0
        for n, info in zip(tb_names, tb_info):
            name = n.get("DSPDISPNAME", "") if isinstance(n, dict) else ""
            if not name:
                continue
            dr = abs(fv(info.get("DSPCLDRAMT", {}).get("DSPCLDRAMTA") if isinstance(info, dict) else 0))
            cr = abs(fv(info.get("DSPCLCRAMT", {}).get("DSPCLCRAMTA") if isinstance(info, dict) else 0))
            is_group = name in PRIMARY_GROUPS
            if is_group:
                total_dr += dr
                total_cr += cr
            rows.append({
                "name":      name,
                "dr":        round(dr, 2),
                "cr":        round(cr, 2),
                "net":       round(cr - dr, 2),
                "is_group":  is_group,
                # clickable = non-group rows with a balance → can drill into vouchers
                "clickable": not is_group and (dr > 0 or cr > 0),
            })

        return jsonify({
            "company":  company,
            "fy_start": fy_start,
            "fy_end":   fy_end,
            "total_dr": round(total_dr, 2),
            "total_cr": round(total_cr, 2),
            "rows":     rows,
        })

    # ── /financials/snapshot  — full parsed financials (BS + P&L + TB) as JSON
    @app.route("/financials/snapshot")
    def financials_snapshot():
        """
        Live fetch from Tally → parse_data → JSON snapshot.
        Returns the exact structure that renderTB / renderPNL / renderBS in
        index.html expect.  No DB required — every call hits Tally live.

        Query params:
          company       (required)
          as_of_date    YYYY-MM-DD or YYYY (year-end)
          from_date     override period start
          to_date       override period end
        """
        company = _require("company")
        today   = date.today()
        fy_year = today.year if today.month >= 4 else today.year - 1

        as_of = _param("as_of_date", str(date(fy_year + 1, 3, 31)))
        try:
            if len(as_of) == 4:
                year     = int(as_of)
                fy_end   = str(date(year, 3, 31))
                fy_start = str(date(year - 1, 4, 1))
            else:
                fy_end   = as_of
                dt_end   = datetime.strptime(as_of, "%Y-%m-%d").date()
                fy_start = str(date(
                    dt_end.year - 1 if dt_end.month <= 3 else dt_end.year, 4, 1))
        except Exception:
            fy_end   = str(date(fy_year + 1, 3, 31))
            fy_start = str(date(fy_year, 4, 1))

        fy_start = _param("from_date", fy_start)
        fy_end   = _param("to_date",   fy_end)

        try:
            dt_end = datetime.strptime(fy_end, "%Y-%m-%d").date()
            # fetch_all wants the year-end year (e.g. 2026 for FY25-26)
            year = dt_end.year if dt_end.month >= 4 else dt_end.year + 1
            raw_data, fs, fe = fetch_all(
                company, year, save_json=False,
                from_date=fy_start, to_date=fy_end
            )
            D = parse_data(company, raw_data, fs, fe)
        except Exception as e:
            return jsonify({"error": str(e)}), 500

        # ── helpers ──────────────────────────────────────────────────────────
        def r2(v):
            try:
                return round(float(v), 2)
            except Exception:
                return 0.0

        def lines_obj(pairs):
            """Convert [(name, value), …] → [{name, amount}] for the frontend."""
            return [{"name": n, "amount": r2(v)} for n, v in pairs if v != 0]

        def lines_signed(pairs):
            """Preserve TB sign so the frontend can decide DR/CR side itself."""
            return [{"name": n, "amount": r2(v)} for n, v in pairs if v != 0]

        # ── Trial Balance ─────────────────────────────────────────────────────
        # renderTB first tries data.trial_balance as a raw Tally object
        # (DSPACCNAME / DSPACCINFO).  We send the pre-processed rows array as
        # fallback (Strategy 2 in renderTB).
        tb_rows = []
        for name in D["TB_ORDER"]:
            val = D["TB"].get(name, 0)
            # Tally TB: DR = negative (asset/expense side), CR = positive
            dr = round(abs(val), 2) if val < 0 else 0.0
            cr = round(val,     2) if val > 0 else 0.0
            tb_rows.append({"name": name, "dr": dr, "cr": cr})

        # ── Dynamic company-type detection ──────────────────────────────────
        # Determines has_stock so that service/trading companies like Vianci
        # never show phantom closing stock in P&L or BS Current Assets.
        _ctype = detect_company_type(D["TB"])

        _open_stock  = r2(D["opening_stock_abs"]) if _ctype["has_stock"] else 0.0
        _close_stock = r2(D["bs_closing_stock"])  if _ctype["has_stock"] else 0.0

        # Gross profit recomputed for each company type:
        #   Manufacturing/Trading WITH stock : full formula including opening/closing stock
        #   Trading WITHOUT stock            : Sales + Purchases + Dir Inc + Dir Exp (no stock terms)
        #   Service / NGO                    : Sales + Dir Inc + Dir Exp only (no purchases)
        if _ctype["has_stock"]:
            _gp = r2(D["gross_profit"])
        elif _ctype["type"] == "trading":
            # Trading company without stock ledgers (e.g. Vianci) — purchases ARE direct costs
            _gp = r2(D["sales"] + D["purchases"] + D["dir_inc"] + D["dir_exp"])
        else:
            # Pure service / NGO — no purchase accounts at all
            # GP = Sales (CR+) + Direct Incomes (CR+) + Direct Expenses (DR-)
            _gp = r2(D["sales"] + D["dir_inc"] + D["dir_exp"])

        # Net profit and BS P&L MUST use the same corrected GP.
        # D["net_profit"] uses raw gross_profit from base tally_core calculation which
        # may include phantom stock terms → wrong for Vianci / service companies.
        # Recompute here so BS P&L "Current Period" matches P&L "Nett Profit" exactly.
        _np          = r2(_gp + D["ind_inc"] + D["ind_exp"])
        _pnl_on_bs   = r2(D["pnl_bf"] + _np)
        _pnl_current = _np

        # ── P&L payload ───────────────────────────────────────────────────────
        pnl_payload = {
            # raw TB-signed (positive = CR, negative = DR) — renderPNL uses sign
            "sales":             r2(D["sales"]),
            "purchases":         r2(D["purchases"]),
            "direct_incomes":    r2(D["dir_inc"]),
            "direct_expenses":   r2(D["dir_exp"]),
            "indirect_incomes":  r2(D["ind_inc"]),
            "indirect_expenses": r2(D["ind_exp"]),
            "opening_stock":     _open_stock,
            "closing_stock":     _close_stock,
            "gross_profit":      _gp,
            "net_profit":        _np,          # ✅ corrected — consistent with _gp
            "pnl_bf":            r2(D["pnl_bf"]),
            # company type flags — used by frontend to hide Trading Account
            "company_type":      _ctype["type"],
            "has_stock":         _ctype["has_stock"],
            # child lines as [{name, amount}]
            "sales_lines":        lines_obj(D["sales_lines"]),
            "purchase_lines":     lines_obj(D["purch_lines"]),
            "direct_inc_lines":   lines_obj(D["dir_inc_lines"]),
            "direct_exp_lines":   lines_obj(D["dir_exp_lines"]),
            "indirect_inc_lines": lines_obj(D["ind_inc_lines"]),
            "indirect_exp_lines": lines_obj(D["ind_exp_lines"]),
            # stock breakdown (empty for non-stock companies)
            "stock_lines": [
                {"name": sl["name"],
                 "closing": r2(sl["closing"]),
                 "opening": r2(sl["opening"])}
                for sl in D["stk_ledgers"]
            ] if _ctype["has_stock"] else [],
            # previous year
            "py_gross_profit":  r2(D["gross_profit_py"]),
            "py_net_profit":    r2(D["net_profit_py"]),
            "py_sales":         r2(D["sales_py"]),
            "py_purchases":     r2(D["purch_py"]),
            "py_closing_stock": r2(D["closing_stock_py"]) if _ctype["has_stock"] else 0.0,
        }

        # ── Balance Sheet ─────────────────────────────────────────────────────
        def _section(total_val, child_pairs, abs_total=False):
            """Build one BS section dict."""
            t = r2(abs(total_val) if abs_total else total_val)
            return {"total": t, "lines": lines_signed(child_pairs)}

        bs_liabilities = {
            "capital_account":    _section(D["cap_total"],   D["cap_lines"]),
            "profit_and_loss":    {
                # ✅ Use corrected _pnl_on_bs / _pnl_current (consistent with _gp/_np)
                # D["pnl_on_bs"] / D["pnl_current"] use raw net_profit which may include
                # phantom stock terms for no-stock companies like Vianci
                "total":   _pnl_on_bs,
                "opening": r2(D["pnl_opening"]),
                "current": _pnl_current,
                "lines": [
                    {"name": "Opening Balance", "amount": r2(D["pnl_opening"])},
                    {"name": "Current Period",  "amount": _pnl_current},
                ],
            },
            "loans":               _section(D["loans_total"], D["loans_lines"]),
            "current_liabilities": _section(D["cl_total"],    D["cl_lines"]),
        }

        # For non-stock companies: remove closing stock line from Current Assets
        _ca_lines_clean = D["ca_lines"]
        _ca_total_clean = D["ca_total"]
        _ca_close_stock = _close_stock
        if not _ctype["has_stock"]:
            _ca_lines_clean = [(n, v) for n, v in D["ca_lines"]
                               if "stock" not in n.lower()]
            # Recalc CA total without the phantom stock
            _ca_total_clean = sum(v for _, v in _ca_lines_clean)
            _ca_close_stock = 0.0

        bs_assets = {
            "fixed_assets":   _section(D["fa_total"], D["fa_lines_all"]),
            "current_assets": {
                "total":         r2(_ca_total_clean),
                "lines":         lines_signed(_ca_lines_clean),
                "closing_stock": r2(_ca_close_stock),
                "stock_lines": [
                    {"name": sl["name"], "amount": r2(sl["closing"])}
                    for sl in D["stk_ledgers"] if sl["closing"] > 0
                ] if _ctype["has_stock"] else [],
            },
        }

        # BS totals (for KPI row in renderBS)
        # ✅ Use corrected _pnl_on_bs so total matches Tally
        total_liab   = (abs(D["cap_total"]) + _pnl_on_bs +
                        abs(D["loans_total"]) + abs(D["cl_total"]))
        total_assets = abs(D["fa_total"]) + abs(_ca_total_clean)

        balance_sheet = {
            "total_liabilities": r2(total_liab),
            "total_assets":      r2(total_assets),
            "liabilities":       bs_liabilities,
            "assets":            bs_assets,
            # previous year group totals
            "py": {
                "capital_total":   r2(abs(D["cap_total_py"])),
                "loans_total":     r2(abs(D["loans_total_py"])),
                "cl_total":        r2(abs(D["cl_total_py"])),
                "fa_total":        r2(abs(D["fa_total_py"])),
                "ca_total":        r2(abs(D["ca_total_py"])),
                "pnl_on_bs":       r2(D["pnl_on_bs_py"]),
                "closing_stock":   r2(D["closing_stock_py"]),
            },
        }

        # ── Ratios ────────────────────────────────────────────────────────────
        sales_abs = abs(D["sales"]) or 1
        ratios = {
            "gross_margin_pct": r2(_gp / sales_abs * 100),   # ✅ corrected GP
            "net_margin_pct":   r2(_np / sales_abs * 100),   # ✅ corrected NP
            "current_ratio":    r2(abs(D["ca_total"]) / abs(D["cl_total"])) if D["cl_total"] else 0,
            "debt_equity":      r2(abs(D["loans_total"]) / abs(D["cap_total"])) if D["cap_total"] else 0,
        }

        # ── Company info ──────────────────────────────────────────────────────
        co = D.get("co_info", {}) or {}
        co_info = {
            "name":    co.get("NAME", company) or company,
            "gstin":   co.get("GSTREGISTRATIONNO", co.get("GSTREGNO", "")) or "",
            "pan":     co.get("PANITNO", "") or "",
            "cin":     co.get("CINNUMBER", co.get("CINNO", "")) or "",
            "state":   co.get("STATENAME", "") or "",
            "address": co.get("ADDRESS", co.get("MAILINGNAME", "")) or "",
        }

        snapshot = {
            "company":       company,
            "fy_start":      fs,
            "fy_end":        fe,
            "as_of_date":    fy_end,
            "co_info":       co_info,
            "trial_balance": tb_rows,
            "pnl":           pnl_payload,
            "balance_sheet": balance_sheet,
            "ratios":        ratios,
            "debtors":       [
                {"name": d.get("name",""), "parent": d.get("parent",""),
                 "balance": r2(d.get("balance",0))}
                for d in (D.get("debtors_list") or [])
            ],
            "creditors":     [
                {"name": d.get("name",""), "parent": d.get("parent",""),
                 "balance": r2(d.get("balance",0))}
                for d in (D.get("creditors_list") or [])
            ],
            "stock_ledgers": [
                {"name": sl["name"], "closing": r2(sl["closing"]),
                 "opening": r2(sl["opening"])}
                for sl in (D.get("stk_ledgers") or [])
            ],
        }

        return jsonify(snapshot)

    # ── /financials/generate  — trigger a fresh fetch + optional PDF save ──
    @app.route("/financials/generate")
    def financials_generate():
        """
        Fetches from Tally and returns the same JSON as /financials/snapshot.
        Also generates a PDF if ?save=true is passed.
        Query: company, from_date, to_date, as_of_date, save (true/false)
        """
        company  = _require("company")
        today    = date.today()
        fy_year  = today.year if today.month >= 4 else today.year - 1
        fy_start = _param("from_date", str(date(fy_year, 4, 1)))
        fy_end   = _param("to_date",   str(date(fy_year + 1, 3, 31)))
        save_pdf = _param("save", "false").lower() == "true"

        try:
            dt_end = datetime.strptime(fy_end, "%Y-%m-%d").date()
            year   = dt_end.year if dt_end.month >= 4 else dt_end.year + 1
            raw_data, fs, fe = fetch_all(
                company, year, save_json=False,
                from_date=fy_start, to_date=fy_end
            )
            D = parse_data(company, raw_data, fs, fe)
        except Exception as e:
            return jsonify({"error": str(e)}), 500

        pdf_path = None
        if save_pdf:
            try:
                import re as _re
                safe_co = _re.sub(r"[^\w]", "_", company[:30])
                pdf_path = f"sch3_{safe_co}_{fy_end}.pdf"
                global STYLES
                STYLES = mk_styles()
                generate_pdf(D, pdf_path)
            except Exception as e:
                pdf_path = None
                print(f"  PDF generation failed: {e}")

        return jsonify({
            "status":    "ok",
            "company":   company,
            "fy_start":  fs,
            "fy_end":    fe,
            "pdf_saved": pdf_path,
            "net_profit":   _np,   # ✅ corrected — consistent with company type
            "gross_profit": _gp,   # ✅ corrected — consistent with company type
            "message":   f"Data fetched live from Tally.{' PDF saved: '+pdf_path if pdf_path else ''}",
        })

    # ── /api/v1/financials  — alias used by older UI calls ────────────────────
    @app.route("/api/v1/financials")
    def api_financials():
        return financials_snapshot()

    # ── /api/v1/chat/context  — Rich AI context with full TB + transactions ──
    @app.route("/api/v1/chat/context")
    def api_chat_context():
        """
        Builds a rich financial context string for the AI chat.
        Includes:
          1. Every TB row with DR/CR amounts (group + ledger level)
          2. Company type detection (service / trading / manufacturing)
          3. Full P&L summary computed from TB, with correct GP for each type
          4. Transaction-level voucher detail for the period (date, type,
             ledger, amount, narration) — enables AI to answer
             "show April sales" or "list all payments to X" precisely.

        Query: company (required), from_date YYYY-MM-DD, to_date YYYY-MM-DD
        """
        company  = _require("company")
        today    = date.today()
        fy_year  = today.year if today.month >= 4 else today.year - 1
        fy_start = _param("from_date", str(date(fy_year, 4, 1)))
        fy_end   = _param("to_date",   str(date(fy_year + 1, 3, 31)))

        parts = [f"Company: {company}", f"Period: {fy_start} to {fy_end}"]

        # ── 1. Full Trial Balance ────────────────────────────────────────────
        try:
            raw_tb = fetch_trial_balance(company, fy_start, fy_end)

            def _fk(obj, key, depth=0):
                if depth > 8 or not isinstance(obj, dict): return None
                if key in obj: return obj[key]
                for v in obj.values():
                    r = _fk(v, key, depth+1)
                    if r is not None: return r
                return None

            def fv2(s):
                if s is None: return 0.0
                try: return float(str(s).replace(",", ""))
                except: return 0.0

            tb_names2 = _fk(raw_tb, "DSPACCNAME") or []
            tb_info2  = _fk(raw_tb, "DSPACCINFO") or []
            if not isinstance(tb_names2, list): tb_names2 = [tb_names2]
            if not isinstance(tb_info2,  list): tb_info2  = [tb_info2]

            TB2 = {}
            tb_ctx_lines = []
            for n, info in zip(tb_names2, tb_info2):
                name = n.get("DSPDISPNAME", "") if isinstance(n, dict) else ""
                if not name: continue
                dr = abs(fv2(info.get("DSPCLDRAMT", {}).get("DSPCLDRAMTA") if isinstance(info, dict) else 0))
                cr = abs(fv2(info.get("DSPCLCRAMT", {}).get("DSPCLCRAMTA") if isinstance(info, dict) else 0))
                TB2[name] = cr - dr   # positive=CR, negative=DR
                if dr > 0 or cr > 0:
                    side = f"DR ₹{dr:,.2f}" if dr >= cr else f"CR ₹{cr:,.2f}"
                    tb_ctx_lines.append(f"  {name}: {side}")

            if tb_ctx_lines:
                parts.append(f"\n=== TRIAL BALANCE ({fy_start} to {fy_end}) ===")
                parts.extend(tb_ctx_lines)

            # ── 2. Company type + P&L summary from TB ───────────────────────
            ctype2   = detect_company_type(TB2)
            sales2   = abs(TB2.get("Sales Accounts", 0))
            purch2   = abs(TB2.get("Purchase Accounts", 0))
            dir_exp2 = abs(TB2.get("Direct Expenses", 0))
            dir_inc2 = abs(TB2.get("Direct Incomes", 0))
            ind_exp2 = abs(TB2.get("Indirect Expenses", 0))
            ind_inc2 = abs(TB2.get("Indirect Incomes", 0))
            open_s2  = abs(TB2.get("Opening Stock", 0)) if ctype2["has_stock"] else 0
            close_s2 = ctype2["closing_stock"]

            if ctype2["has_stock"]:
                gp2 = sales2 + dir_inc2 - purch2 - dir_exp2 - open_s2 + close_s2
            else:
                gp2 = sales2 + dir_inc2 - purch2 - dir_exp2
            np2 = gp2 + ind_inc2 - ind_exp2

            parts.append(f"\nCompany Type: {ctype2['type']} | Has Stock: {ctype2['has_stock']}")
            parts.append(f"\n=== P&L SUMMARY ===")
            parts.append(f"  Sales:             ₹{sales2:,.2f}")
            parts.append(f"  Purchases:         ₹{purch2:,.2f}")
            parts.append(f"  Direct Expenses:   ₹{dir_exp2:,.2f}")
            if ctype2["has_stock"]:
                parts.append(f"  Opening Stock:     ₹{open_s2:,.2f}")
                parts.append(f"  Closing Stock:     ₹{close_s2:,.2f}")
            parts.append(f"  Gross Profit:      ₹{gp2:,.2f}  (GP% {(gp2/sales2*100 if sales2 else 0):.1f}%)")
            parts.append(f"  Indirect Expenses: ₹{ind_exp2:,.2f}")
            parts.append(f"  Indirect Incomes:  ₹{ind_inc2:,.2f}")
            parts.append(f"  Net Profit:        ₹{np2:,.2f}  (NP% {(np2/sales2*100 if sales2 else 0):.1f}%)")

        except Exception as e:
            parts.append(f"\n[TB error: {e}]")

        # ── 3. Transaction-level voucher detail for the period ───────────────
        try:
            vouchers = fetch_vouchers_from_tally(company, fy_start, fy_end)

            # Type-level summary
            vtype_totals: dict = {}
            for v in vouchers:
                vt = v.get("type", "Other") or "Other"
                amt_v = v.get("total_dr", 0) or v.get("total_cr", 0)
                if vt not in vtype_totals:
                    vtype_totals[vt] = {"count": 0, "amount": 0.0}
                vtype_totals[vt]["count"]  += 1
                vtype_totals[vt]["amount"] += float(amt_v)

            parts.append(f"\n=== TRANSACTIONS ({fy_start} to {fy_end}) ===")
            parts.append(f"  Total vouchers: {len(vouchers)}")
            for vt, d in sorted(vtype_totals.items(), key=lambda x: -x[1]["amount"]):
                parts.append(f"  {vt}: {d['count']} vouchers | ₹{d['amount']:,.2f}")

            # Entry-level rows — up to 300 (most recent first for AI relevance)
            _limit = 300
            parts.append(f"\n--- TRANSACTION ENTRIES (latest {min(len(vouchers), _limit)} vouchers) ---")
            for v in vouchers[-_limit:]:
                vdate  = v.get("date", "")
                vtype  = v.get("type", "")
                vnum   = v.get("number", "")
                narr   = (v.get("narration") or "")[:60]
                party  = v.get("party", "")
                for e in v.get("entries", []):
                    ldg   = e.get("ledger", "")
                    eamt  = e.get("amount", 0)
                    etype = e.get("type", "")   # "DR" or "CR"
                    parts.append(
                        f"  {vdate} | {vtype} {vnum} | {ldg} | {etype} ₹{eamt:,.2f}"
                        + (f" | Party: {party}" if party and party != ldg else "")
                        + (f" | {narr}" if narr else "")
                    )
        except Exception as e:
            parts.append(f"\n[Voucher fetch error: {e}]")

        ctx_text = "\n".join(parts)
        return jsonify({
            "ok":           True,
            "company":      company,
            "fy_start":     fy_start,
            "fy_end":       fy_end,
            "context":      ctx_text,
            "context_text": ctx_text,   # alias for index.html compatibility
        })

    # ── /api/v1/chat/sessions  (stub — in-memory sessions) ──────────────────
    _chat_sessions = {}   # session_key -> [{role, content}]

    @app.route("/api/v1/chat/sessions", methods=["GET"])
    def api_chat_sessions_list():
        company = _param("company", "")
        sessions = [
            {"id": k, "company": v.get("company", ""), "messages": len(v.get("history", []))}
            for k, v in _chat_sessions.items()
            if not company or v.get("company", "") == company
        ]
        return jsonify({"sessions": sessions})

    @app.route("/api/v1/chat/sessions", methods=["POST"])
    def api_chat_sessions_create():
        from flask import request as freq2
        body    = freq2.get_json(silent=True) or {}
        company = body.get("company", "")
        sid     = f"s{len(_chat_sessions)+1}_{company[:10].replace(' ','_')}"
        _chat_sessions[sid] = {"company": company, "history": []}
        return jsonify({"id": sid, "company": company})

    @app.route("/api/v1/chat/sessions/<sid>", methods=["GET", "DELETE"])
    def api_chat_session(sid):
        from flask import request as freq3
        if freq3.method == "DELETE":
            _chat_sessions.pop(sid, None)
            return jsonify({"deleted": sid})
        sess = _chat_sessions.get(sid, {})
        return jsonify({"id": sid, "history": sess.get("history", [])})

    # ── /api/v1/stored-reports  (file-based — saves generated PDF paths) ───
    import glob as _glob

    @app.route("/api/v1/stored-reports")
    def api_stored_reports():
        company = _param("company", "")
        pattern = f"sch3_{'*' if not company else company[:30].replace(' ','_') + '*'}.pdf"
        files   = sorted(_glob.glob(pattern), reverse=True)
        reports = []
        for f in files:
            try:
                sz = os.path.getsize(f)
                mt = datetime.fromtimestamp(os.path.getmtime(f)).isoformat()
            except Exception:
                sz, mt = 0, ""
            reports.append({"id": f, "filename": f, "size": sz, "created_at": mt})
        return jsonify({"reports": reports, "total": len(reports)})

    @app.route("/api/v1/stored-report/<path:report_id>", methods=["GET", "DELETE"])
    def api_stored_report(report_id):
        from flask import request as freq4, send_file as _sf
        if freq4.method == "DELETE":
            try:
                os.remove(report_id)
                return jsonify({"deleted": report_id})
            except Exception as e:
                return jsonify({"error": str(e)}), 404
        if os.path.exists(report_id):
            return _sf(os.path.abspath(report_id), mimetype="application/pdf")
        return jsonify({"error": "Not found"}), 404

    # ── /api/v1/ledger/vouchers  — all vouchers touching one ledger ──────────
    @app.route("/api/v1/ledger/vouchers")
    def api_ledger_vouchers():
        """
        Returns every voucher that posts to a specific ledger — identical to
        Tally's Ledger Voucher report.
        Query: company, ledger (ledger name), from_date, to_date
        """
        company = _require("company")
        ledger  = _require("ledger")
        today   = date.today()
        fy_year = today.year if today.month >= 4 else today.year - 1
        from_date = _param("from_date", str(date(fy_year, 4, 1)))
        to_date   = _param("to_date",   str(today))
        try:
            vouchers = fetch_ledger_vouchers(company, ledger, from_date, to_date)
            total_dr = sum(v["total_dr"] for v in vouchers)
            total_cr = sum(v["total_cr"] for v in vouchers)
            return jsonify({
                "company":    company,
                "ledger":     ledger,
                "from_date":  from_date,
                "to_date":    to_date,
                "total":      len(vouchers),
                "total_dr":   round(total_dr, 2),
                "total_cr":   round(total_cr, 2),
                "vouchers":   vouchers,
            })
        except Exception as e:
            return jsonify({"error": str(e), "vouchers": []}), 500

    # ── /api/v1/vouchers/by-type  — all vouchers of one type ─────────────────
    @app.route("/api/v1/vouchers/by-type")
    def api_vouchers_by_type():
        """
        Returns all vouchers of a given type. Tries Collection API first,
        falls back to day-chunk fetch.
        Query: company, type (voucher type name), from_date, to_date
        """
        company  = _require("company")
        vtype    = _require("type")
        today    = date.today()
        fy_year  = today.year if today.month >= 4 else today.year - 1
        from_date = _param("from_date", str(date(fy_year, 4, 1)))
        to_date   = _param("to_date",   str(today))
        try:
            vouchers = fetch_vouchers_by_type(company, vtype, from_date, to_date)
            total_dr = sum(v["total_dr"] for v in vouchers)
            total_cr = sum(v["total_cr"] for v in vouchers)
            return jsonify({
                "company":    company,
                "type":       vtype,
                "from_date":  from_date,
                "to_date":    to_date,
                "total":      len(vouchers),
                "total_dr":   round(total_dr, 2),
                "total_cr":   round(total_cr, 2),
                "vouchers":   vouchers,
            })
        except Exception as e:
            return jsonify({"error": str(e), "vouchers": []}), 500

    # ── /api/v1/voucher/<guid>  — single voucher full detail ─────────────────
    @app.route("/api/v1/voucher/<guid>")
    def api_voucher_detail(guid):
        """
        Fetch one specific voucher by GUID from Tally.
        Query: company
        """
        company = _require("company")
        xml_body = (
            f"<ENVELOPE><HEADER><VERSION>1</VERSION>"
            f"<TALLYREQUEST>Export</TALLYREQUEST>"
            f"<TYPE>Object</TYPE><SUBTYPE>Voucher</SUBTYPE>"
            f"<ID>{xe(guid)}</ID></HEADER>"
            f"<BODY><DESC>"
            f"<STATICVARIABLES>"
            f"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
            f"<SVCURRENTCOMPANY>{xe(company)}</SVCURRENTCOMPANY>"
            f"</STATICVARIABLES>"
            f"<FETCH>*</FETCH>"
            f"</DESC></BODY></ENVELOPE>"
        )
        try:
            raw      = post(xml_body, timeout=30)
            vouchers = _parse_voucher_xml(raw)
            if vouchers:
                return jsonify({"voucher": vouchers[0]})
            return jsonify({"error": "Voucher not found"}), 404
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    # ── /api/v1/tb/ledger_vouchers  — TB drill-down: vouchers for one ledger ──
    @app.route("/api/v1/tb/ledger_vouchers")
    def api_tb_ledger_vouchers():
        """
        Returns every voucher that touches a specific ledger — powers the
        click-to-expand drill-down panel in the Trial Balance tab.

        Query: company, ledger, from_date, to_date
        Response: {ledger, total_dr, total_cr, total, vouchers:[
            {date, type, number, narration, amount, dr_cr,
             party, other_ledgers, is_cancelled}
        ]}
        """
        company   = _require("company")
        ledger    = _require("ledger")
        today     = date.today()
        fy_year   = today.year if today.month >= 4 else today.year - 1
        from_date = _param("from_date", str(date(fy_year, 4, 1)))
        to_date   = _param("to_date",   str(today))

        try:
            raw_vouchers = fetch_ledger_vouchers(company, ledger, from_date, to_date)
        except Exception as e:
            return jsonify({"error": str(e), "vouchers": [], "total": 0}), 500

        out       = []
        total_dr  = 0.0
        total_cr  = 0.0

        for v in raw_vouchers:
            entries  = v.get("entries", [])
            # Find this ledger's own entry
            my_entry = next(
                (e for e in entries if e.get("ledger", "").lower() == ledger.lower()),
                entries[0] if entries else None
            )

            amt     = float(my_entry.get("amount", 0)) if my_entry else 0.0
            etype   = (my_entry.get("type", "") if my_entry else "") or ""
            # If type field not populated, derive from IsDeemedPositive logic
            if not etype:
                etype = "DR" if amt > 0 else "CR"

            amt_abs = abs(amt)
            if etype == "DR":
                total_dr += amt_abs
            else:
                total_cr += amt_abs

            # Other ledgers involved (contra side) — for narration context
            others = [
                e.get("ledger", "")
                for e in entries
                if e.get("ledger", "").lower() != ledger.lower() and e.get("ledger")
            ][:5]

            out.append({
                "date":          v.get("date", ""),
                "type":          v.get("type", ""),
                "number":        v.get("number", ""),
                "narration":     (v.get("narration") or "")[:120],
                "amount":        round(amt_abs, 2),
                "dr_cr":         etype,
                "party":         v.get("party", ""),
                "other_ledgers": others,
                "is_cancelled":  v.get("is_cancelled", False),
            })

        out.sort(key=lambda x: x["date"])

        return jsonify({
            "company":   company,
            "ledger":    ledger,
            "from_date": from_date,
            "to_date":   to_date,
            "total":     len(out),
            "total_dr":  round(total_dr, 2),
            "total_cr":  round(total_cr, 2),
            "vouchers":  out,
        })

    # ── /api/v1/sch3/classic  — Generate Classic Schedule III PDF ───────────
    @app.route("/api/v1/sch3/classic")
    def api_sch3_classic():
        """
        Fetches data from Tally and generates a Classic Schedule III PDF
        using sch3_classic.py.

        Query params:
          company    (required) — exact Tally company name
          from_date  YYYY-MM-DD — period start (default: current FY start)
          to_date    YYYY-MM-DD — period end   (default: current FY end)
          download   true/false  — stream PDF to browser (default: true)

        Custom date range example:
          /api/v1/sch3/classic?company=Vianci+Enterprise&from_date=2026-04-01&to_date=2026-07-31

        Full FY example:
          /api/v1/sch3/classic?company=Vianci+Enterprise&from_date=2026-04-01&to_date=2027-03-31
        """
        from flask import send_file
        import io, tempfile, os as _os

        company  = _require("company")
        today    = date.today()
        fy_year  = today.year if today.month >= 4 else today.year - 1
        from_date = _param("from_date", str(date(fy_year, 4, 1)))
        to_date   = _param("to_date",   str(date(fy_year + 1, 3, 31)))
        download  = _param("download", "true").lower() != "false"

        # Validate dates
        try:
            _fd = datetime.strptime(from_date, "%Y-%m-%d").date()
            _td = datetime.strptime(to_date,   "%Y-%m-%d").date()
            if _fd >= _td:
                return jsonify({"error": "from_date must be before to_date"}), 400
        except ValueError:
            return jsonify({"error": "Dates must be YYYY-MM-DD format"}), 400

        year = _td.year if _td.month >= 4 else _td.year + 1

        try:
            # 1. Fetch from Tally
            raw_data, fs, fe = fetch_all(
                company, year, save_json=False,
                from_date=from_date, to_date=to_date
            )
            D = parse_data(company, raw_data, fs, fe)
        except Exception as e:
            return jsonify({"error": f"Tally fetch failed: {e}"}), 500

        try:
            # 2. Generate Classic Schedule III PDF via sch3_classic.py
            import sch3_classic as s3c

            # Previous year D (auto-fetch for PY column)
            D_py = None
            try:
                _py_end   = str(date(_fd.year - 1 if _fd.month <= 3 else _fd.year, 3, 31))
                _py_start = str(date(date.fromisoformat(_py_end).year - 1, 4, 1))
                raw_py, fs_py, fe_py = fetch_all(
                    company, date.fromisoformat(_py_end).year,
                    save_json=False,
                    from_date=_py_start, to_date=_py_end
                )
                D_py_candidate = parse_data(company, raw_py, fs_py, fe_py)
                if not s3c._py_is_same_as_cy(D, D_py_candidate):
                    D_py = D_py_candidate
            except Exception as py_err:
                print(f"  [PY fetch skipped: {py_err}]")

            # 3. Write to temp file and stream back
            safe_co  = re.sub(r'[^\w]', '_', company[:30])
            filename = f"classic_{safe_co}_{to_date}.pdf"

            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp_path = tmp.name

            s3c.generate_classic_pdf(D, D_py, tmp_path)

            if download:
                with open(tmp_path, "rb") as f:
                    pdf_bytes = f.read()
                _os.unlink(tmp_path)
                return send_file(
                    io.BytesIO(pdf_bytes),
                    mimetype="application/pdf",
                    as_attachment=True,
                    download_name=filename
                )
            else:
                # Save alongside server and return path
                dest = filename
                import shutil
                shutil.move(tmp_path, dest)
                return jsonify({
                    "ok":       True,
                    "company":  company,
                    "from":     from_date,
                    "to":       to_date,
                    "filename": dest,
                    "message":  f"Classic Schedule III saved → {dest}"
                })

        except ImportError:
            return jsonify({
                "error": "sch3_classic.py not found. Place it in the same folder as tally_core.py"
            }), 500
        except Exception as e:
            return jsonify({"error": f"PDF generation failed: {e}"}), 500

    return app


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="Tally → Schedule III PDF  (one command, live fetch)",
        epilog="""Examples:
  python tally_sch3.py --list-companies
  python tally_sch3.py --company "DEEP DEMO" --year 2026
  python tally_sch3.py --company "DEEP DEMO" --year 2026 --open
  python tally_sch3.py --company "DEEP DEMO" --year 2026 --port 9001 --out my_report.pdf""",
        formatter_class=argparse.RawDescriptionHelpFormatter)

    ap.add_argument("--company",        help='Exact company name as in Tally')
    ap.add_argument("--year",  type=int,help='Financial year END year e.g. 2026 = FY 2025-26')
    ap.add_argument("--from-date",      help='Custom period start YYYY-MM-DD (overrides --year)')
    ap.add_argument("--to-date",        help='Custom period end   YYYY-MM-DD (overrides --year)')
    ap.add_argument("--port",  type=int,default=9000, help='Tally HTTP port (default 9000)')
    ap.add_argument("--out",            help='Output PDF filename (auto-generated if omitted)')
    ap.add_argument("--open", action="store_true", help='Auto-open PDF after generation')
    ap.add_argument("--list-companies", action="store_true", help='List all companies open in Tally')
    ap.add_argument("--no-json",        action="store_true", help='Skip saving raw JSON file')
    ap.add_argument("--serve", action="store_true",
                    help='Start Flask server (serves index.html + live voucher API on port 5057)')
    ap.add_argument("--server-port", type=int, default=5057,
                    help='Port for Flask server (default 5057)')
    args = ap.parse_args()

    TALLY_URL = f"http://localhost:{args.port}"

    # ── serve mode ───────────────────────────────────────────────────────────
    if args.serve:
        set_tally_url(f"http://localhost:{args.port}")
        flask_app = _make_flask_server(port=args.server_port, tally_port=args.port)
        print(f"\n{'═'*65}")
        print(f"  TallySync Live Server")
        print(f"  Tally   : http://localhost:{args.port}")
        print(f"  Browser : http://localhost:{args.server_port}")
        print(f"  Open    : http://localhost:{args.server_port}  in your browser")
        print(f"{'═'*65}")
        print(f"  Endpoints:")
        print(f"    GET /daybook?company=X&from_date=YYYY-MM-DD&to_date=YYYY-MM-DD")
        print(f"    GET /api/v1/vouchers?company=X&from=YYYY-MM-DD&to=YYYY-MM-DD")
        print(f"    GET /api/v1/summary?company=X")
        print(f"    GET /voucher_types?company=X")
        print(f"    GET /companies")
        print(f"{'═'*65}\n")
        flask_app.run(host="0.0.0.0", port=args.server_port, debug=False, threaded=True)
        sys.exit(0)

    # ── list companies ────────────────────────────────────────────────────────
    if args.list_companies:
        print(f"\n  Connecting to Tally at {TALLY_URL} ...")
        try: req.post(TALLY_URL, data=b"<ENVELOPE/>", timeout=4)
        except:
            print(f"  ERROR: Tally not reachable at {TALLY_URL}"); sys.exit(1)
        cos = list_companies()
        if cos:
            print(f"\n  {len(cos)} compan{'y' if len(cos)==1 else 'ies'} found in Tally:\n")
            for i, c in enumerate(cos, 1):
                print(f"    {i:>3}.  {c}")
            print(f'\n  Usage:  python tally_sch3.py --company "NAME" --year 2026\n')
        else:
            print("  No companies found. Is Tally open with a company loaded?")
        sys.exit(0)

    # ── validate args ─────────────────────────────────────────────────────────
    if not args.company:
        ap.error('--company "COMPANY NAME" is required  (or use --list-companies)')

    # Resolve date range: custom dates take priority over --year
    _from_date = getattr(args, 'from_date', None)
    _to_date   = getattr(args, 'to_date',   None)

    if _from_date and _to_date:
        # Custom date range — validate format
        try:
            _fd = datetime.strptime(_from_date, "%Y-%m-%d").date()
            _td = datetime.strptime(_to_date,   "%Y-%m-%d").date()
            if _fd >= _td:
                ap.error("--from-date must be earlier than --to-date")
        except ValueError:
            ap.error("Dates must be in YYYY-MM-DD format  e.g. --from-date 2026-04-01")
        # Derive year from to_date for fetch_all internal use
        _year = _td.year if _td.month >= 4 else _td.year + 1
        print(f"\n  Custom date range: {_from_date}  →  {_to_date}")
    else:
        if not args.year:
            ap.error('--year YYYY is required  e.g. --year 2026  for FY 2025-26\n'
                     'Or use --from-date YYYY-MM-DD --to-date YYYY-MM-DD for a custom range')
        if args.year < 2000 or args.year > 2100:
            ap.error('--year must be a 4-digit year e.g. 2026')
        _year      = args.year
        _from_date = None
        _to_date   = None

    # ── fetch → parse → PDF ──────────────────────────────────────────────────
    raw_data, fy_start, fy_end = fetch_all(
        args.company, _year, save_json=not args.no_json,
        from_date=_from_date, to_date=_to_date)

    print("\n  Parsing data...", end=" ", flush=True)
    D = parse_data(args.company, raw_data, fy_start, fy_end)
    print(f"OK  ({len(D['TB'])} TB rows, {sum(len(v) for v in D['stock_by_cat'].values())} stock items)")

    safe    = re.sub(r'[^\w]','_', args.company[:30])
    out     = args.out or f"sch3_{safe}_{fy_end}.pdf"

    print(f"\n  Building PDF...", end=" ", flush=True)
    generate_pdf(D, out)

    print(f"\n{'═'*65}")
    print(f"  Done!  Report: {out}")
    print(f"  FY   : {fy_start}  to  {fy_end}")
    print(f"  GP   : ₹ {fmt_indian(abs(D['gross_profit']))}")
    print(f"  NP   : ₹ {fmt_indian(abs(D['net_profit']))}  ({'Loss' if D['net_profit']<0 else 'Profit'})")
    print(f"{'═'*65}\n")

    if args.open:
        try:
            if sys.platform == "win32":    os.startfile(out)
            elif sys.platform == "darwin": subprocess.run(["open", out])
            else:                          subprocess.run(["xdg-open", out])
        except: pass