# -*- coding: utf-8 -*-
"""
sch3_verify.py — Verification Summary + Statutory Disclosures
=============================================================================
Add-on module for sch3_classic.py.
Builds three additional pages appended after accounting policies:

  PAGE A — Verification Summary
            BS balance check, P&L tie, company type, share capital match,
            prior year status — visible ticks/flags for CA to sign off

  PAGE B — MSME Creditor Disclosure
            Schedule III 2021 amendment — mandatory statutory requirement
            Auto-pulled from Tally creditors list

  PAGE C — Related Party Disclosure (AS-18)
            Auto-detects director/partner names vs loan ledger names
            Flags probable related party transactions

Imported and called from generate_classic_pdf() in sch3_classic.py.
"""

import re
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (Table, TableStyle, Paragraph, Spacer, PageBreak,
                                 HRFlowable)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

import tally_core as tc

PAGE_W = tc.PAGE_W
C_BLK  = colors.black
C_GREEN  = colors.HexColor("#1A7A1A")
C_RED    = colors.HexColor("#C00000")
C_ORANGE = colors.HexColor("#C47A00")
C_BLUE   = colors.HexColor("#1F3864")
C_LGREY  = colors.HexColor("#F2F2F2")
C_DGREY  = colors.HexColor("#D5D5D5")

# ── Styles ───────────────────────────────────────────────────────────────────
def _styles():
    return {
        "TITLE": ParagraphStyle("VT", fontName="Helvetica-Bold", fontSize=11,
                                 alignment=TA_CENTER, leading=14, textColor=C_BLUE),
        "SUB":   ParagraphStyle("VS", fontName="Helvetica-Bold", fontSize=9,
                                 alignment=TA_CENTER, leading=11, textColor=C_BLUE),
        "HDR":   ParagraphStyle("VH", fontName="Helvetica-Bold", fontSize=8,
                                 alignment=TA_LEFT,  leading=10, textColor=colors.white),
        "NRM":   ParagraphStyle("VN", fontName="Helvetica",      fontSize=8,
                                 alignment=TA_LEFT,  leading=10),
        "NRM_R": ParagraphStyle("VR", fontName="Helvetica",      fontSize=8,
                                 alignment=TA_RIGHT, leading=10),
        "BOLD":  ParagraphStyle("VB", fontName="Helvetica-Bold", fontSize=8,
                                 alignment=TA_LEFT,  leading=10),
        "SML":   ParagraphStyle("VM", fontName="Helvetica",      fontSize=7,
                                 alignment=TA_LEFT,  leading=9,  textColor=colors.HexColor("#555555")),
        "SML_R": ParagraphStyle("VMR",fontName="Helvetica",      fontSize=7,
                                 alignment=TA_RIGHT, leading=9,  textColor=colors.HexColor("#555555")),
        "OK":    ParagraphStyle("VOK",fontName="Helvetica-Bold", fontSize=9,
                                 alignment=TA_CENTER,leading=11, textColor=C_GREEN),
        "WARN":  ParagraphStyle("VW", fontName="Helvetica-Bold", fontSize=9,
                                 alignment=TA_CENTER,leading=11, textColor=C_RED),
        "INFO":  ParagraphStyle("VI", fontName="Helvetica",      fontSize=8,
                                 alignment=TA_LEFT,  leading=10, textColor=C_ORANGE),
    }

S = _styles()

def P(t, s="NRM"):  return Paragraph(str(t) if t else "", S[s])
def fmt(v):
    try:
        v = float(v)
        if abs(v) < 0.005: return "—"
        s = tc.fmt_indian(abs(v))
        return f"({s})" if v < 0 else s
    except: return "—"

def fmt_date(d):
    s = str(d or "").strip()
    if len(s) == 10 and s[4] == "-":
        return f"{s[8:10]}.{s[5:7]}.{s[:4]}"
    return s

def _grid(rows, cw, shaded=(), bold_last=False):
    ts = [
        ("GRID",         (0,0), (-1,-1), 0.5, C_BLK),
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",   (0,0), (-1,-1), 2.5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 2.5),
        ("LEFTPADDING",  (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("BACKGROUND",   (0,0), (-1,0),  C_BLUE),
        ("FONTNAME",     (0,0), (-1,0),  "Helvetica-Bold"),
        ("TEXTCOLOR",    (0,0), (-1,0),  colors.white),
    ]
    for r in shaded:
        ts.append(("BACKGROUND", (0,r), (-1,r), C_LGREY))
    if bold_last:
        ts += [
            ("BACKGROUND",  (0,-1), (-1,-1), C_DGREY),
            ("FONTNAME",    (0,-1), (-1,-1), "Helvetica-Bold"),
        ]
    return Table(rows, colWidths=cw, style=TableStyle(ts))

# ══════════════════════════════════════════════════════════════════════════════
# PAGE A — VERIFICATION SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
def build_verification_page(D, D_py, story, bs_f=None):
    """
    Adds a Verification Summary page to story.
    Shows pass/fail for all key financial integrity checks.
    Visible in the PDF so the CA sees it before signing.

    bs_f: pre-computed bs_figures(D) from generate_classic_pdf.
          Uses it directly so verification matches the actual BS page.
          Falls back to recomputing if not supplied (backward-compatible).
    """
    from sch3_classic import bs_figures, pnl_figures

    story.append(PageBreak())
    co   = D.get("company", "")
    fy_e = D.get("fy_end", "")

    story.append(P(co.upper(), "TITLE"))
    story.append(P(f"SCHEDULE III — VERIFICATION SUMMARY", "SUB"))
    story.append(P(f"For the year ended {fmt_date(fy_e)}  |  Generated: {datetime.now().strftime('%d.%m.%Y %H:%M')}", "SML"))
    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_BLUE, spaceAfter=4))

    checks = []   # list of (description, status, detail)  status: OK/WARN/INFO/FAIL

    # ── 1. Company type detected ──────────────────────────────────────────────
    TB = D.get("TB", {})
    co_type = tc.detect_company_type(TB)
    ctype_label = co_type.get("type", "unknown").upper()
    is_dormant  = co_type.get("is_dormant", False)
    checks.append((
        "Company Type Detected",
        "INFO",
        f"{ctype_label}{'  (Dormant — no transactions this year)' if is_dormant else ''}",
    ))

    # ── 2. Tally connection & data completeness ───────────────────────────────
    tb_rows = len(TB)
    checks.append((
        "Trial Balance Rows Fetched",
        "OK" if tb_rows > 5 else "WARN",
        f"{tb_rows} rows  {'✓ Adequate' if tb_rows > 5 else '⚠ Very few rows — verify Tally connection'}",
    ))

    # ── 3. Prior year availability ────────────────────────────────────────────
    if D_py:
        py_fy = D_py.get("fy_end", "")
        checks.append(("Prior Year Data", "OK", f"Available — FY ending {fmt_date(py_fy)} ✓"))
    else:
        checks.append(("Prior Year Data", "INFO",
                        "Not available — prior year column shows '—'  "
                        "(acceptable for first year of operations)"))

    # ── 4. Balance Sheet balance check ───────────────────────────────────────
    # BUG 10 FIX: use pre-computed bs_f so verification matches the actual BS page
    f = bs_f if bs_f is not None else bs_figures(D)
    total_l = f.get("total_l", 0) if f else 0
    total_a = f.get("total_a", 0) if f else 0
    bs_diff = abs(total_a - total_l)
    if bs_diff < 1.0:
        checks.append(("Balance Sheet", "OK",
                        f"BALANCES ✓  Total Assets = Total Liabilities = {fmt(total_l)}"))
    elif bs_diff < 500:
        checks.append(("Balance Sheet", "WARN",
                        f"Rounding difference Rs.{bs_diff:.2f}  |  Assets: {fmt(total_a)}  "
                        f"Liabilities: {fmt(total_l)}  — Review before signing"))
    else:
        checks.append(("Balance Sheet", "FAIL",
                        f"DOES NOT BALANCE  |  Assets: {fmt(total_a)}  "
                        f"Liabilities: {fmt(total_l)}  |  Difference: {fmt(bs_diff)}  "
                        f"— DO NOT SIGN until resolved"))

    # ── 5. Share capital check ────────────────────────────────────────────────
    eq_share  = D.get("equity_share", 0)
    cap_total = abs(D.get("cap_total", 0))
    _SHARE_KW_V = ('share capital', 'equity share', 'preference share', 'equity capital')
    _cap_lines_v = D.get('cap_lines', [])
    _has_share_cap_v = any(any(kw in nm.lower() for kw in _SHARE_KW_V) for nm, _ in _cap_lines_v)
    if eq_share > 0:
        checks.append(("Share Capital", "OK",
                        f"Rs.{fmt(eq_share)} extracted from {len(_cap_lines_v)} "
                        f"capital ledger(s) ✓"))
    elif not _has_share_cap_v and _cap_lines_v is not None:
        checks.append(("Share Capital", "INFO",
                        "LLP / Proprietorship / Partnership — no share capital. "
                        "Note 1 shows Partners' / Proprietor's Capital ✓"))
    else:
        checks.append(("Share Capital", "WARN",
                        "Could not detect share capital from ledger names — "
                        "verify Note 1 manually"))

    # ── 6. P&L tie to TB ─────────────────────────────────────────────────────
    np_computed = D.get("net_profit", 0)
    # Sign convention: TB['P&L A/c'] negative = DR = loss (use directly, no negation)
    _tb_pnl_v   = D.get("TB", {}).get("Profit & Loss A/c", 0)
    pnl_on_bs   = D.get("pnl_on_bs", 0)
    pnl_bf      = D.get("pnl_bf", 0)
    # If pnl_on_bs was wrongly negated (abs match with opposite sign), restore it
    if _tb_pnl_v != 0 and abs(pnl_on_bs + _tb_pnl_v) < 1.0:
        pnl_on_bs = _tb_pnl_v
        pnl_bf    = _tb_pnl_v
    pnl_check   = abs((pnl_bf + np_computed) - pnl_on_bs)
    if pnl_check < 1.0:
        if is_dormant:
            checks.append(("P&L Tie to TB", "OK",
                            f"Dormant year — Net Profit: NIL  "
                            f"|  Accumulated Loss B/F: {fmt(pnl_bf)} ✓"))
        else:
            checks.append(("P&L Tie to TB", "OK",
                            f"Net Profit: {fmt(np_computed)}  "
                            f"|  P&L on BS: {fmt(pnl_on_bs)} ✓"))
    else:
        checks.append(("P&L Tie to TB", "FAIL",
                        f"P&L mismatch — NP: {fmt(np_computed)}  "
                        f"BF: {fmt(pnl_bf)}  Sum: {fmt(pnl_bf+np_computed)}  "
                        f"BS shows: {fmt(pnl_on_bs)}  — Investigate"))

    # ── 7. Closing stock check ────────────────────────────────────────────────
    if co_type.get("has_stock"):
        cs = D.get("bs_closing_stock", 0)
        n_stk = len(D.get("stk_ledgers", []))
        checks.append(("Closing Stock", "OK",
                        f"{fmt(cs)} from {n_stk} stock ledger(s) ✓"))
    else:
        checks.append(("Closing Stock", "INFO",
                        "No inventory — stock terms set to NIL ✓"))

    # ── 8. Fixed assets check ─────────────────────────────────────────────────
    fa_net = D.get("fa_net_bs", 0)
    fa_gross = D.get("fa_gross", 0)
    dep = D.get("dep_reserve", 0)
    if fa_gross > 0:
        checks.append(("Fixed Assets", "OK",
                        f"Gross Block: {fmt(fa_gross)}  "
                        f"Less Dep: {fmt(dep)}  Net: {fmt(fa_net)} ✓"))
    else:
        checks.append(("Fixed Assets", "INFO", "No fixed assets recorded in Tally"))

    # ── 9. Unsecured loans disclosure check ──────────────────────────────────
    loans_lines = D.get("loans_lines", [])
    unsec = [(n,v) for n,v in loans_lines if "unsecured" in n.lower()]
    if unsec:
        total_unsec = sum(abs(v) for _,v in unsec)
        checks.append(("Unsecured Loans", "WARN",
                        f"Rs.{fmt(total_unsec)} — Verify related party disclosure "
                        f"(lender name, rate, relationship) in Note 3"))
    else:
        checks.append(("Unsecured Loans", "OK", "No unsecured loans ✓"))

    # ── 10. MSME disclosure readiness ────────────────────────────────────────
    creditors = D.get("creditors_list", [])
    checks.append(("MSME Creditor Disclosure", "INFO",
                    f"{len(creditors)} creditor(s) found — "
                    "MSME declaration included in next page  "
                    "(verify with client whether any are MSME-registered)"))

    # ── Render checks table ───────────────────────────────────────────────────
    STATUS_COLOR = {
        "OK":   colors.HexColor("#E8F5E9"),
        "WARN": colors.HexColor("#FFF8E1"),
        "FAIL": colors.HexColor("#FFEBEE"),
        "INFO": colors.HexColor("#E3F2FD"),
    }
    STATUS_ICON = {"OK": "✓  PASS", "WARN": "⚠  REVIEW", "FAIL": "✗  FAIL", "INFO": "ℹ  INFO"}

    CW = [PAGE_W*0.30, PAGE_W*0.15, PAGE_W*0.55]
    rows = [[P("Check", "HDR"), P("Status", "HDR"), P("Detail", "HDR")]]
    ts_extra = []
    for i, (desc, status, detail) in enumerate(checks, 1):
        icon = STATUS_ICON.get(status, status)
        icon_color = {"OK": C_GREEN, "WARN": C_ORANGE, "FAIL": C_RED, "INFO": C_BLUE}[status]
        icon_style = ParagraphStyle(f"IC{i}", parent=S["BOLD"], fontSize=7.5,
                                     textColor=icon_color, alignment=TA_CENTER)
        rows.append([
            Paragraph(desc, S["NRM"]),
            Paragraph(icon, icon_style),
            Paragraph(detail, S["SML"]),
        ])
        ts_extra.append(("BACKGROUND", (0,i), (-1,i), STATUS_COLOR.get(status, C_LGREY)))

    tbl = Table(rows, colWidths=CW, style=TableStyle([
        ("GRID",          (0,0), (-1,-1), 0.5, colors.HexColor("#AAAAAA")),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 5),
        ("RIGHTPADDING",  (0,0), (-1,-1), 5),
        ("BACKGROUND",    (0,0), (-1,0),  C_BLUE),
        ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
        ("TEXTCOLOR",     (0,0), (-1,0),  colors.white),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [C_LGREY, colors.white]),
    ] + ts_extra))
    story.append(tbl)
    story.append(Spacer(1, 4*mm))

    # ── Overall verdict ───────────────────────────────────────────────────────
    fail_count = sum(1 for _,s,_ in checks if s == "FAIL")
    warn_count = sum(1 for _,s,_ in checks if s == "WARN")
    if fail_count == 0 and warn_count == 0:
        verdict_text = "ALL CHECKS PASSED — Financial statements appear complete and consistent."
        verdict_color = C_GREEN
        verdict_bg    = colors.HexColor("#E8F5E9")
    elif fail_count == 0:
        verdict_text = (f"{warn_count} item(s) require review before signing. "
                        "See ⚠ rows above.")
        verdict_color = C_ORANGE
        verdict_bg    = colors.HexColor("#FFF8E1")
    else:
        verdict_text = (f"{fail_count} CRITICAL issue(s) detected. "
                        "DO NOT SIGN until resolved. See ✗ rows above.")
        verdict_color = C_RED
        verdict_bg    = colors.HexColor("#FFEBEE")

    verdict_style = ParagraphStyle("VRD", fontName="Helvetica-Bold", fontSize=9,
                                    alignment=TA_CENTER, textColor=verdict_color,
                                    leading=13)
    vt = Table([[Paragraph(verdict_text, verdict_style)]],
               colWidths=[PAGE_W],
               style=TableStyle([
                   ("BACKGROUND", (0,0), (-1,-1), verdict_bg),
                   ("GRID",       (0,0), (-1,-1), 1.0, verdict_color),
                   ("TOPPADDING", (0,0), (-1,-1), 6),
                   ("BOTTOMPADDING",(0,0),(-1,-1),6),
               ]))
    story.append(vt)
    story.append(Spacer(1, 3*mm))
    story.append(P(
        "Note: This verification summary is auto-generated by the Schedule III tool based on data "
        "fetched from TallyPrime. It does not substitute for professional judgement. The signing "
        "CA/Director should review the financial statements independently before authentication.",
        "SML"
    ))


# ══════════════════════════════════════════════════════════════════════════════
# PAGE B — MSME CREDITOR DISCLOSURE
# ══════════════════════════════════════════════════════════════════════════════
def build_msme_disclosure(D, D_py, story):
    """
    Schedule III Amendment 2021 — mandatory MSME creditor disclosure.
    Companies must disclose dues to Micro and Small Enterprises separately.
    Auto-pulls creditor list from Tally. CA must verify MSME status with client.
    """
    story.append(PageBreak())
    co   = D.get("company", "")
    fy_e = D.get("fy_end", "")

    story.append(P(co.upper(), "TITLE"))
    story.append(P("NOTE 5A — MSME CREDITOR DISCLOSURE", "SUB"))
    story.append(P(
        f"Pursuant to Schedule III Amendment (MCA Notification dated 24th March 2021) "
        f"for the year ended {fmt_date(fy_e)}",
        "SML"
    ))
    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_BLUE, spaceAfter=4))

    # ── Statutory disclosure table ────────────────────────────────────────────
    # These are the 6 mandatory disclosures under Schedule III Amendment 2021
    fy_p = D_py.get("fy_end", "") if D_py else ""
    CW = [PAGE_W*0.70, PAGE_W*0.15, PAGE_W*0.15]

    rows = [[
        P("Particulars", "HDR"),
        Paragraph(f"<b>{fmt_date(fy_e)}</b>", S["HDR"]),
        Paragraph(f"<b>{fmt_date(fy_p) if fy_p else '—'}</b>", S["HDR"]),
    ]]
    msme_items = [
        ("(i)   Principal amount remaining unpaid as at year end",                        None, None),
        ("(ii)  Interest due on above remaining unpaid as at year end",                   None, None),
        ("(iii) Amount of interest paid in terms of Section 16 of MSMED Act, 2006 "
         "along with the payment made beyond appointed date",                             None, None),
        ("(iv)  Amount of interest due and payable for the period of delay in payment "
         "(other than interest specified under MSMED Act, 2006)",                         None, None),
        ("(v)   Amount of interest accrued and remaining unpaid as at year end",          None, None),
        ("(vi)  Amount of further interest remaining due and payable in succeeding year", None, None),
    ]
    for i, (label, cy, py) in enumerate(msme_items):
        bg = C_LGREY if i % 2 == 0 else colors.white
        rows.append([
            Paragraph(label, S["NRM"]),
            Paragraph("—", S["NRM_R"]),
            Paragraph("—", S["NRM_R"]),
        ])

    story.append(_grid(rows, CW, shaded=tuple(range(1, len(rows), 2))))
    story.append(Spacer(1, 4*mm))

    # ── Creditor list for CA reference ───────────────────────────────────────
    creditors = D.get("creditors_list", [])
    story.append(P("Creditors as per TallyPrime — MSME Status to be Confirmed by Client", "BOLD"))
    story.append(Spacer(1, 2*mm))

    if creditors:
        CW2 = [PAGE_W*0.05, PAGE_W*0.55, PAGE_W*0.25, PAGE_W*0.15]
        crows = [[P("#","HDR"), P("Creditor Name","HDR"),
                  P("Balance (Rs.)","HDR"), P("MSME?","HDR")]]
        for i, c in enumerate(creditors, 1):
            name = c.get("name", "")
            bal  = c.get("balance", 0)
            crows.append([
                P(str(i), "SML"),
                P(name, "SML"),
                Paragraph(fmt(abs(bal)), S["SML_R"]),
                P("Verify with client", "SML"),
            ])
        story.append(_grid(crows, CW2, shaded=tuple(range(1, len(crows), 2))))
    else:
        story.append(P("No creditors found in Tally for this period.", "NRM"))

    story.append(Spacer(1, 4*mm))

    # ── Declaration ───────────────────────────────────────────────────────────
    decl_style = ParagraphStyle("DECL", fontName="Helvetica", fontSize=7.5,
                                  alignment=TA_LEFT, leading=11,
                                  textColor=colors.HexColor("#333333"))
    story.append(Table([[Paragraph(
        "<b>Declaration by Management:</b> Based on information and explanations received from "
        "the Company, none of the creditors of the Company are registered under the Micro, Small "
        "and Medium Enterprises Development Act, 2006 (MSMED Act). Accordingly, the disclosures "
        "required under Schedule III to the Companies Act, 2013 with respect to amounts due to "
        "Micro and Small Enterprises are not applicable / are NIL. "
        "<br/><br/>"
        "<i>Note to CA: Please confirm the above declaration with the client management before "
        "finalising the financial statements. If any creditor is MSME-registered, update the "
        "above table with actual figures.</i>",
        decl_style
    )]], colWidths=[PAGE_W],
    style=TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), colors.HexColor("#FFFDE7")),
        ("GRID",          (0,0),(-1,-1), 0.8, C_ORANGE),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ("RIGHTPADDING",  (0,0),(-1,-1), 8),
    ])))


# ══════════════════════════════════════════════════════════════════════════════
# PAGE C — RELATED PARTY DISCLOSURE (AS-18)
# ══════════════════════════════════════════════════════════════════════════════
def _detect_related_parties(D):
    """
    Auto-detect probable related party transactions by comparing:
    - Director / partner names from Capital Account ledger names
    - Loan ledger names from Loans (Liability) group
    - Creditor names from creditors_list
    Returns list of dicts: {name, relationship, ledger, amount, type}
    """
    related = []

    # Extract person names from capital account ledger names
    cap_lines  = D.get("cap_lines", [])
    loans_lines = D.get("loans_lines", [])
    creditors   = D.get("creditors_list", [])

    # Build list of probable director/partner names from cap_lines
    # Strip common suffixes: "Share Capital", "Capital A/c", "Current A/c", "Drawing A/c"
    STRIP_RX = re.compile(
        r'\b(share\s+capital|capital\s+a[/.]?c|current\s+a[/.]?c|'
        r'drawing\s+a[/.]?c|drawings|account|a/c|pvt|ltd|llp|private|limited)\b',
        re.IGNORECASE
    )
    person_names = []
    _RP_SKIP_KW = ('mediclaim', 'school fee', 'tution fee', 'tuition fee',
                   'insurance', 'rent', 'salary', ' exp', 'expense', 'fees',
                   'telephone', 'electric', 'petrol', 'donation', 'profit',
                   'p&l', 'surplus', 'reserve', 'share capital', 'equity')
    for n, v in cap_lines:
        nl = n.lower()
        if any(kw in nl for kw in _RP_SKIP_KW):
            continue
        if nl.startswith('drawing'):
            continue
        cleaned = STRIP_RX.sub("", n).strip(" -/.,")
        if cleaned and len(cleaned) > 3:
            # Use cleaned name as display name (not raw ledger name with suffixes)
            person_names.append((cleaned.lower(), cleaned, abs(v)))

    # Match person names against loan ledger names
    for pname_low, pname_orig, cap_val in person_names:
        for lname, lval in loans_lines:
            if pname_low in lname.lower() or any(
                tok in lname.lower() for tok in pname_low.split() if len(tok) > 3
            ):
                related.append({
                    "name":         pname_orig,
                    "relationship": "Director / Partner",
                    "ledger":       lname,
                    "amount":       abs(lval),
                    "type":         "Loan from Related Party",
                    "note":         "Unsecured loan — confirm interest rate and repayment terms",
                })

    # Match person names against creditor names
    for pname_low, pname_orig, cap_val in person_names:
        for c in creditors:
            cname = c.get("name", "")
            if pname_low in cname.lower() or any(
                tok in cname.lower() for tok in pname_low.split() if len(tok) > 3
            ):
                related.append({
                    "name":         pname_orig,
                    "relationship": "Director / Partner",
                    "ledger":       cname,
                    "amount":       abs(c.get("balance", 0)),
                    "type":         "Creditor (Related Party)",
                    "note":         "Trade payable to related party — confirm arm's length",
                })

    return related


def build_related_party_disclosure(D, D_py, story):
    """
    AS-18 Related Party Disclosure — auto-detected from Tally data.
    """
    story.append(PageBreak())
    co   = D.get("company", "")
    fy_e = D.get("fy_end", "")

    story.append(P(co.upper(), "TITLE"))
    story.append(P("NOTE 31A — RELATED PARTY DISCLOSURES (AS-18)", "SUB"))
    story.append(P(
        f"As required by Accounting Standard 18 — Related Party Disclosures  "
        f"|  For the year ended {fmt_date(fy_e)}",
        "SML"
    ))
    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_BLUE, spaceAfter=4))

    # ── A. Key Management Personnel ───────────────────────────────────────────
    story.append(P("A. Key Management Personnel (KMP) and Related Parties Identified", "BOLD"))
    story.append(Spacer(1, 2*mm))

    cap_lines = D.get("cap_lines", [])
    STRIP_RX  = re.compile(
        r'\b(share\s+capital|capital\s+a[/.]?c|current\s+a[/.]?c|'
        r'drawing\s+a[/.]?c|drawings|account|a/c|pvt|ltd|llp|private|limited)\b',
        re.IGNORECASE
    )

    kmp_list = []
    _KMP_SKIP_KW = ('mediclaim', 'school fee', 'tution fee', 'tuition fee',
                    'insurance', 'rent', 'salary', ' exp', 'expense', 'fees',
                    'telephone', 'electric', 'petrol', 'donation', 'profit',
                    'p&l', 'surplus', 'reserve', 'share capital', 'equity')
    for n, v in cap_lines:
        nl = n.lower()
        if any(kw in nl for kw in _KMP_SKIP_KW):
            continue
        if nl.startswith('drawing'):
            continue
        cleaned = STRIP_RX.sub("", n).strip(" -/.,")
        if cleaned and len(cleaned) > 3:
            kmp_list.append((cleaned, "Director / Partner / Proprietor"))

    if kmp_list:
        CW_KMP = [PAGE_W*0.05, PAGE_W*0.55, PAGE_W*0.40]
        krows  = [[P("#","HDR"), P("Name","HDR"), P("Designation / Relationship","HDR")]]
        for i, (name, role) in enumerate(kmp_list, 1):
            krows.append([P(str(i),"SML"), P(name,"NRM"), P(role,"NRM")])
        story.append(_grid(krows, CW_KMP, shaded=tuple(range(1,len(krows),2))))
    else:
        story.append(P("No KMP identified from capital account ledger names. "
                        "Please update manually.", "INFO"))

    story.append(Spacer(1, 4*mm))

    # ── B. Related Party Transactions ─────────────────────────────────────────
    story.append(P("B. Related Party Transactions Auto-Detected from Tally", "BOLD"))
    story.append(Spacer(1, 2*mm))

    related = _detect_related_parties(D)

    if related:
        CW_T = [PAGE_W*0.22, PAGE_W*0.18, PAGE_W*0.22, PAGE_W*0.15, PAGE_W*0.23]
        trows = [[P("Party Name","HDR"), P("Relationship","HDR"),
                  P("Ledger in Tally","HDR"), P("Amount (Rs.)","HDR"), P("Note","HDR")]]
        for r in related:
            trows.append([
                P(r["name"],         "NRM"),
                P(r["relationship"], "NRM"),
                P(r["ledger"],       "NRM"),
                Paragraph(fmt(r["amount"]), S["NRM_R"]),
                P(r["note"],         "SML"),
            ])
        story.append(_grid(trows, CW_T, shaded=tuple(range(1,len(trows),2))))
        story.append(Spacer(1,2*mm))
        story.append(P(
            "⚠  Above transactions are auto-detected based on name matching. "
            "Please verify with management and add any transactions not listed above.",
            "INFO"
        ))
    else:
        story.append(Table([[Paragraph(
            "No related party transactions auto-detected from Tally data.\n\n"
            "Based on information and explanations received from the management, "
            "there were no transactions with related parties as defined under AS-18 "
            "during the year ended " + fmt_date(fy_e) + ".",
            S["NRM"]
        )]], colWidths=[PAGE_W],
        style=TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), colors.HexColor("#E8F5E9")),
            ("GRID",          (0,0),(-1,-1), 0.8, C_GREEN),
            ("TOPPADDING",    (0,0),(-1,-1), 6),
            ("BOTTOMPADDING", (0,0),(-1,-1), 6),
            ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ])))

    story.append(Spacer(1, 4*mm))

    # ── C. Standard declaration ───────────────────────────────────────────────
    story.append(P("C. Declaration", "BOLD"))
    story.append(Spacer(1, 2*mm))
    story.append(Table([[Paragraph(
        "The above related party relationships and transactions have been identified on "
        "the basis of information available with the Company and the representations made "
        "by the key management personnel, which have been relied upon by the auditors. "
        "The transactions were carried out in the ordinary course of business and at "
        "arm's length unless otherwise stated.",
        S["NRM"]
    )]], colWidths=[PAGE_W],
    style=TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), C_LGREY),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ])))


# ══════════════════════════════════════════════════════════════════════════════
# UPDATED ACCOUNTING POLICIES — company-type aware
# ══════════════════════════════════════════════════════════════════════════════
def build_accounting_policies(D, story):
    """
    Company-type aware accounting policies.
    Auto-fills company name, FY dates. Suppresses irrelevant policies.
    Replaces tc.build_policies() for Schedule III output.
    """
    story.append(PageBreak())
    co    = D.get("company", "")
    fy_e  = D.get("fy_end", "")
    fy_s  = D.get("fy_start", "")
    co_i  = D.get("co_info", {}) or {}
    gstin = co_i.get("GSTREGISTRATIONNO", co_i.get("GSTREGNO", "")) or ""
    pan   = co_i.get("PANITNO", "") or ""
    cin   = co_i.get("CINNUMBER", co_i.get("CINNO", "")) or ""
    state = co_i.get("STATENAME", "") or ""

    TB      = D.get("TB", {})
    co_type = tc.detect_company_type(TB)
    has_stock   = co_type.get("has_stock", False)
    has_fa      = D.get("fa_gross", 0) > 0
    has_forex   = abs(TB.get("Unadjusted Forex Gain/Loss",
                      TB.get("Unadjusted Forex Gain / Loss", 0))) > 0
    has_exp     = abs(TB.get("Indirect Expenses", 0)) > 0 or abs(TB.get("Direct Expenses", 0)) > 0
    is_dormant  = co_type.get("is_dormant", False)

    # Detect entity type for correct legal citation in Basis of Preparation
    _cap_lines_pol = D.get('cap_lines', [])
    _SHARE_KW_POL  = ('share capital', 'equity share', 'preference share', 'equity capital')
    _has_share_pol = any(any(kw in nm.lower() for kw in _SHARE_KW_POL)
                        for nm, _ in _cap_lines_pol)
    co_name_l = co.lower()
    _is_llp  = "llp" in co_name_l or "limited liability" in co_name_l
    _is_prop = (not _has_share_pol and not _is_llp)  # proprietorship or partnership
    _is_company = _has_share_pol and not _is_llp

    # Header
    HDR_S = ParagraphStyle("PH", fontName="Helvetica-Bold", fontSize=11,
                             alignment=TA_CENTER, textColor=C_BLUE, leading=14)
    SUB_S = ParagraphStyle("PS", fontName="Helvetica-Bold", fontSize=9,
                             alignment=TA_CENTER, textColor=C_BLUE, leading=12)
    NRM_S = ParagraphStyle("PN", fontName="Helvetica", fontSize=8,
                             alignment=TA_LEFT,  leading=11)
    BLD_S = ParagraphStyle("PB", fontName="Helvetica-Bold", fontSize=8,
                             alignment=TA_LEFT,  leading=11)

    story.append(Paragraph(co.upper(), HDR_S))
    story.append(Paragraph("SIGNIFICANT ACCOUNTING POLICIES", SUB_S))
    story.append(Paragraph(
        f"For the year ended {fmt_date(fy_e)}"
        + (f"  |  GSTIN: {gstin}" if gstin else "")
        + (f"  |  PAN: {pan}" if pan else "")
        + (f"  |  CIN: {cin}" if cin else ""),
        ParagraphStyle("PI", fontName="Helvetica", fontSize=7, alignment=TA_CENTER,
                        textColor=colors.HexColor("#555555"), leading=10)
    ))
    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width=PAGE_W, thickness=1.5, color=C_BLUE, spaceAfter=5))

    # ── Policies — company-type aware ────────────────────────────────────────
    policies = []

    if _is_company:
        _basis_text = (
            f"The financial statements of {co} have been prepared in accordance with the Generally "
            "Accepted Accounting Principles in India (Indian GAAP) and comply with the Accounting "
            "Standards specified under Section 133 of the Companies Act, 2013 read with Rule 7 of "
            "the Companies (Accounts) Rules, 2014 and the relevant provisions of the Companies "
            f"Act, 2013. The financial statements are prepared on the historical cost convention "
            f"on accrual basis for the period from {fmt_date(fy_s)} to {fmt_date(fy_e)}."
        )
    elif _is_llp:
        _basis_text = (
            f"The financial statements of {co} have been prepared in accordance with the Generally "
            "Accepted Accounting Principles in India (Indian GAAP) and the Accounting Standards "
            "issued by the Institute of Chartered Accountants of India (ICAI), to the extent "
            "applicable, and in accordance with the Limited Liability Partnership Act, 2008 and "
            "the rules made thereunder. The financial statements are prepared on the historical "
            f"cost convention on accrual basis for the period from {fmt_date(fy_s)} to {fmt_date(fy_e)}."
        )
    else:
        # Proprietorship / Partnership
        _basis_text = (
            f"The financial statements of {co} have been prepared in accordance with the Generally "
            "Accepted Accounting Principles in India (Indian GAAP) and the Accounting Standards "
            "issued by the Institute of Chartered Accountants of India (ICAI), to the extent "
            "applicable. The financial statements are prepared on the historical cost convention "
            f"on accrual basis for the period from {fmt_date(fy_s)} to {fmt_date(fy_e)}."
        )
    policies.append(("1. Basis of Preparation", _basis_text))

    if not is_dormant:
        policies.append(("2. Revenue Recognition",
            "Revenue from sale of goods is recognised when the significant risks and rewards "
            "of ownership of the goods are transferred to the buyer, which generally coincides "
            "with delivery of goods. Revenue from services is recognised on completion of "
            "services as per the terms of the contract. Interest income is recognised on a "
            "time-proportion basis taking into account the amount outstanding and the rate "
            "applicable."))

    if has_stock:
        policies.append(("3. Inventories",
            "Inventories are valued at lower of cost or net realisable value. Cost of raw "
            "materials and packing materials is determined on weighted average basis. Cost of "
            "work-in-progress and finished goods includes material cost and appropriate share "
            "of manufacturing overheads. Closing stock values are sourced directly from the "
            "TallyPrime inventory register as at the balance sheet date."))

    if has_fa:
        _dep_ref = (
            "Schedule II to the Companies Act, 2013" if _is_company
            else "the rates generally adopted for similar assets / as per Income Tax Act, 1961"
        )
        policies.append(("4. Fixed Assets and Depreciation",
            "Fixed assets are stated at historical cost less accumulated depreciation. Cost "
            "includes purchase price, taxes, freight and other directly attributable costs of "
            "bringing the asset to working condition for intended use. Depreciation is provided "
            f"on the Written Down Value (WDV) method at {_dep_ref}. "
            "Assets costing Rs.5,000 or less are fully depreciated in the year of acquisition."))

    if has_forex:
        policies.append(("5. Foreign Currency Transactions",
            "Foreign currency transactions are recorded at the exchange rate prevailing on the "
            "date of transaction. Monetary assets and liabilities denominated in foreign "
            "currency are restated at the closing rate as at the balance sheet date. Exchange "
            "differences arising on settlement or restatement of monetary items are recognised "
            "in the Statement of Profit and Loss in the period in which they arise."))

    policies.append(("6. Taxation",
        "Current tax is computed and provided for on the basis of taxable income determined "
        "in accordance with the provisions of the Income Tax Act, 1961 and applicable rules "
        "thereunder. Deferred tax is recognised on timing differences between accounting "
        "income and taxable income using the liability method at the tax rates that are "
        "expected to apply when the liability is settled or the asset is realised. Deferred "
        "tax assets are recognised only when there is reasonable certainty of realisation."))

    if has_fa:
        policies.append(("7. Borrowing Costs",
            "Borrowing costs attributable to acquisition or construction of qualifying assets "
            "are capitalised as part of the cost of those assets, up to the date when the "
            "assets are ready for intended use. All other borrowing costs are charged to the "
            "Statement of Profit and Loss in the period in which they are incurred."))

    policies.append(("8. Provisions, Contingent Liabilities and Contingent Assets",
        "Provisions are recognised when the Company has a present obligation (legal or "
        "constructive) as a result of a past event, it is probable that an outflow of "
        "resources embodying economic benefits will be required to settle the obligation "
        "and a reliable estimate can be made of the amount of the obligation. Contingent "
        "liabilities are not recognised but are disclosed in the notes. Contingent assets "
        "are neither recognised nor disclosed."))

    if has_exp:
        policies.append(("9. Employee Benefits",
            "Short-term employee benefit obligations are measured at the undiscounted amount "
            "of the benefits expected to be paid in exchange for the related service. "
            "Provident fund contributions are charged to the Statement of Profit and Loss "
            "as incurred. Gratuity, if applicable, is provided on actuarial basis or as "
            "per Payment of Gratuity Act, 1972, whichever is higher."))

    for title, text in policies:
        story.append(Spacer(1, 2.5*mm))
        story.append(Paragraph(f"<b>{title}</b>", BLD_S))
        story.append(Paragraph(text, NRM_S))

    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width=PAGE_W, thickness=0.8, color=C_DGREY, spaceAfter=4))
    story.append(Paragraph(
        f"For and on behalf of the Board of Directors of {co}  |  "
        f"Place: {state or 'India'}  |  Date: {fmt_date(fy_e)}",
        ParagraphStyle("PF", fontName="Helvetica", fontSize=7.5, alignment=TA_CENTER,
                        textColor=colors.HexColor("#555555"), leading=10)
    ))


# ══════════════════════════════════════════════════════════════════════════════
# MASTER ENTRY POINT
# Called from generate_classic_pdf() after notes section
# ══════════════════════════════════════════════════════════════════════════════
def build_all_statutory(D, D_py, story, bs_f=None):
    """
    Appends all statutory and verification pages to story in order:
      1. Accounting Policies (company-type aware, replaces generic version)
      2. Verification Summary (BS check, P&L tie, warnings)
      3. MSME Creditor Disclosure (Schedule III 2021 amendment)
      4. Related Party Disclosure (AS-18 auto-detected)

    bs_f: pre-computed bs_figures(D) dict from generate_classic_pdf.
          Passed in so the verification page uses the same figures as the BS page.
          If None, bs_figures(D) is recomputed (backward-compatible).
    """
    build_accounting_policies(D, story)
    build_verification_page(D, D_py, story, bs_f=bs_f)
    build_msme_disclosure(D, D_py, story)
    build_related_party_disclosure(D, D_py, story)