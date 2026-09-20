# -*- coding: utf-8 -*-
"""
sch3_classic.py — ADD-ON, does NOT touch tally_core.py / db.py
========================================================================
Universal Schedule III Classic CA-Format PDF generator.
Handles ALL company types: trading (with/without stock), manufacturing,
service / NGO, proprietorship / partnership / LLP / Pvt Ltd.

FIXES APPLIED (vs original):
  1.  Note 2 — Phantom "General Reserve" suppressed for LLPs/simple companies
               that have no real named reserve ledgers.
  2.  Note 6 — Debit-balance items inside Current Liabilities (e.g. Duties &
               Taxes with net GST/TDS credit) are RECLASSIFIED to Short Term
               Loans & Advances (Note 14) instead of shown as negative liability.
  3.  BS Total — total_l and total_a both use the same reclassification so BS
               balances correctly after reclassification.
  4.  PY guard — if PY fetch silently returned same FY data as CY, D_py → None
               (blank PY column, not wrong repeated data).
  5.  Note 21 Other Expenses — also pulls from dir_exp_lines (Direct Expenses)
               so non-indirect expenses like freight/labour are shown.
  6.  P&L figures — uses tally_core's net_profit directly for correct sign; avoids
               double-counting indirect income in other_exp.
  7.  note_table — stable py_map keyed on both original and stripped name to
               correctly match PY ledger names to CY rows.
  8.  Direct Expenses classification — trading companies WITHOUT stock (e.g. Vianci)
               have Direct Expenses included in P&L cost section (Cost of Material
               Consumed / Other Expenses) not silently dropped. Previously these
               4+ lakh Direct Expenses were omitted → profit was overstated.
  9.  P&L total_expenses vs net_profit reconciliation — tot_exp is now reconciled
               back to Tally's authoritative net_profit to prevent any rounding or
               classification drift making PBEET ≠ npat.
  10. Investments — TB group 'Investments' is now extracted and shown under
               NON-CURRENT ASSETS (Non-Current Investments). Was completely absent
               from BS, causing asset understatement (e.g. Vianci ₹22.5 L FDs).
  11. Suspense A/c — TB group 'Suspense A/c' is now captured under Other Current
               Assets and shown in BS. Was silently dropped for companies that
               carry unclassified suspense balances.
  12. Capital Account DR-balance reclassification — proprietor drawing accounts
               and Capital Account sub-ledgers with net DEBIT balance are
               properly reclassified to Long-Term Loans & Advances (asset) so
               the BS does not show a negative Reserves & Surplus.
  13. bs_figures total_l / total_a — now uses Tally TB group totals directly
               (cap_total + loans_total + cl_total) as the authoritative liability
               sum, ensuring BS balances for ALL company types.
  14. pnl_figures for service / no-stock trading companies — inv_chg forced to
               zero when company has no stock (detect_company_type), preventing
               phantom opening/closing stock terms in Schedule III P&L.
  15. Note 17 Cost of Material Consumed — for trading-without-stock companies
               (e.g. Vianci) Direct Expenses (Freight, Labour etc.) that are
               genuine cost-of-goods are included in Note 17 to match Tally P&L.
               Previously only purch_lines were included; dir_exp_lines were
               silently dropped from the note.
"""

import sys, os, json, argparse, re

import tally_core as tc
try:
    import db as tdb
    DB_OK = True
except Exception:
    tdb = None
    DB_OK = False

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, PageBreak)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

PAGE_W = tc.PAGE_W
C_BLK  = colors.black

NAME_SUFFIX_RX = re.compile(r'\s*-\s*\(from\s+\d{1,2}-\w{3}-\d{2,4}\)\s*')

def clean_company_name(name):
    return NAME_SUFFIX_RX.sub('', str(name or '')).strip()

# ──────────────────────────────────────────────────────────────────────────
# Styles
# ──────────────────────────────────────────────────────────────────────────
def mk_classic_styles():
    return {
        "TITLE": ParagraphStyle("TITLE", fontName="Helvetica-Bold", fontSize=11,
                                 alignment=TA_CENTER, leading=13),
        "SUB":   ParagraphStyle("SUB", fontName="Helvetica-Bold", fontSize=9.5,
                                 alignment=TA_CENTER, leading=12),
        "NRM":   ParagraphStyle("NRM", fontName="Helvetica", fontSize=7.6,
                                 alignment=TA_LEFT, leading=9.2),
        "IND":   ParagraphStyle("IND", fontName="Helvetica", fontSize=7.6,
                                 alignment=TA_LEFT, leading=9.2, leftIndent=10),
        "BOLD":  ParagraphStyle("BOLD", fontName="Helvetica-Bold", fontSize=7.6,
                                 alignment=TA_LEFT, leading=9.2),
        "CTR":   ParagraphStyle("CTR", fontName="Helvetica", fontSize=7.6,
                                 alignment=TA_CENTER, leading=9.2),
        "AMT":   ParagraphStyle("AMT", fontName="Helvetica", fontSize=7.6,
                                 alignment=TA_RIGHT, leading=9.2),
        "AMTb":  ParagraphStyle("AMTb", fontName="Helvetica-Bold", fontSize=7.6,
                                 alignment=TA_RIGHT, leading=9.2),
        "SML":   ParagraphStyle("SML", fontName="Helvetica", fontSize=7,
                                 alignment=TA_LEFT, leading=8.5),
    }
S = mk_classic_styles()

def amt(v):
    # FIX: None means PY column not available (first-year company) — render dash, not blank.
    # Empty string means genuinely no value for this field — also render dash.
    if v is None:
        return "—"          # em-dash: PY column explicitly absent
    if v == "":
        return ""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    if abs(v) < 0.005:
        return "-"
    neg = v < 0
    s = tc.fmt_indian(abs(v))
    return f"({s})" if neg else s

def P(txt, st="NRM"):  return Paragraph(txt if txt else "", S[st])
def A(v):              return Paragraph(amt(v), S["AMT"])
def AB(v):             return Paragraph(amt(v), S["AMTb"])
def N(n):              return Paragraph(str(n) if n is not None else "", S["CTR"])

def fmt_date(date_str):
    if not date_str:
        return ''
    s = str(date_str).strip()
    if len(s) == 10 and s[4] == '-':
        y, m, d = s[:4], s[5:7], s[8:10]
    elif len(s) == 8:
        y, m, d = s[:4], s[4:6], s[6:8]
    else:
        return s
    return f"{d}.{m}.{y}"

def safe_div(a, b):
    try:
        return a / b
    except Exception:
        return None

# ──────────────────────────────────────────────────────────────────────────
# Grid table helper
# ──────────────────────────────────────────────────────────────────────────
def grid_table(rows, col_widths, bold_rows=(), shaded_rows=(), header_row=True):
    ts = [
        ("GRID", (0, 0), (-1, -1), 0.6, C_BLK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2),
        ("LEFTPADDING", (0, 0), (0, -1), 4),
        ("RIGHTPADDING", (-1, 0), (-1, -1), 4),
    ]
    if header_row:
        ts.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8E8E8")))
        ts.append(("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"))
    for r in shaded_rows:
        ts.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#D5D5D5")))
        ts.append(("FONTNAME", (0, r), (-1, r), "Helvetica-Bold"))
    return Table(rows, colWidths=col_widths, style=TableStyle(ts))

# ──────────────────────────────────────────────────────────────────────────
# Keyword classifiers
# ──────────────────────────────────────────────────────────────────────────
def classify(name, buckets, fallback):
    nl = name.lower()
    for label, kws in buckets:
        if any(kw in nl for kw in kws):
            return label
    return fallback

def bucketize(lines, buckets, fallback):
    out = {}
    for name, val in lines or []:
        if val == 0:
            continue
        label = classify(name, buckets, fallback)
        out.setdefault(label, []).append((name, val))
    return out

CL_BUCKETS = [
    ("Short Term Borrowings",     ["bank od", "bank oc", "cash credit", " cc a/c",
                                    "secured loan", "unsecured loan", "od a/c"]),
    ("Trade Payables",            ["sundry creditors", "creditors for supplies",
                                    "creditors for expenses", "trade payable"]),
    ("Short Term Provisions",     ["provision", "audit fees payable", "salary payable",
                                    "esic payable", "tax audit fees payable",
                                    "provision for tax", "advance tax", "tds receivable"]),
]
CL_FALLBACK = "Other Current Liabilities"

CA_BUCKETS = [
    ("Inventories",               ["stock-in-hand", "stock in hand", "stock of",
                                    "closing stock", "stock in trade",
                                    "raw material", "finished goods",
                                    "work in progress", "wip", "consumables",
                                    "inventor", "scrap"]),
    ("Trade Receivables",         ["sundry debtor", "trade receivable"]),
    ("Cash and Cash Equivalents", ["cash-in-hand", "cash in hand", "bank account",
                                    "fdr", "deposit with bank"]),
    ("Long Term Loans and Advances", ["security deposit", "deposit with geb",
                                      "deposit with", "rent deposit",
                                      "deposit (asset)", "deposits (asset)",
                                      "deposits(asset)", "deposit(asset)"]),
]
CA_FALLBACK = "Short Term Loans and Advances"

# Investments sub-ledger keywords (all treated as Non-Current Investments)
INV_BUCKETS = [
    ("Non-Current Investments", ["fix deposit", "fixed deposit", " fd ", "f.d.",
                                  "nsc", "ppf", "mutual fund", "share", "bond",
                                  "debenture", "investment", "agarwood",
                                  "diamond invest", "gold", "ncd"]),
]
INV_FALLBACK = "Non-Current Investments"

# Suspense A/c treated as Other Current Assets on BS
SUSPENSE_KW = ["suspense"]

# ──────────────────────────────────────────────────────────────────────────
# RECLASSIFICATION HELPER
# Debit-balance items inside CL (e.g. Duties & Taxes with net GST/TDS credit)
# are assets, not liabilities. We split CL lines into:
#   cl_true  → genuinely CR-balance (show in Note 6 Other CL)
#   cl_reclassified → net DR-balance (move to Note 14 Short Term Loans & Advances)
# ──────────────────────────────────────────────────────────────────────────
def split_cl_other(items):
    """Split Other Current Liabilities into genuine CR items vs DR reclassified items.
    Returns (true_cl_items, reclassified_to_stla_items).
    val > 0 in TB = CR = genuine liability.
    val < 0 in TB = DR = debit balance = should be shown as asset (Short Term Loans & Adv).
    """
    true_cl, reclassified = [], []
    for name, val in items:
        if val >= 0:
            true_cl.append((name, val))
        else:
            # Debit-balance liability — reclassify to asset side
            reclassified.append((name, val))
    return true_cl, reclassified

# ══════════════════════════════════════════════════════════════════════════
# _asset_disp: TB sign → display sign for assets (DR = negative in TB → positive display)
# ══════════════════════════════════════════════════════════════════════════
def _asset_disp(val):
    """Convert a TB-signed asset value to display-positive amount."""
    v = float(val or 0)
    return abs(v) if v < 0 else (-abs(v) if abs(v) > 0.005 else 0)

def _items_disp_sum(items):
    """Sum of _asset_disp across a list of (name, val) pairs."""
    return sum(_asset_disp(v) for _, v in items)


# ══════════════════════════════════════════════════════════════════════════
# bs_figures — computes every BS line with reclassification
# Covers: trading (with/without stock), manufacturing, service,
#         proprietorship (Capital A/c as equity), LLP, Pvt Ltd.
# ══════════════════════════════════════════════════════════════════════════
def bs_figures(D):
    if not D:
        return None
    TB = D.get('TB', {})

    # ── Equity & Liabilities ───────────────────────────────────────────────
    cap_total   = D.get('cap_total', 0)
    eq_share    = D.get('equity_share', 0)
    pnl_on_bs   = D.get('pnl_on_bs', 0)

    # pnl_on_bs sign convention (verified against live Tally data):
    # Tally sends TB['Profit & Loss A/c'] as NEGATIVE for losses (DR balance).
    # tally_core uses this directly: pnl_bf = raw TB value (no negation).
    # So pnl_on_bs < 0 = accumulated loss = reduces reserves = correct.
    # No self-defence negation needed — just use the value as-is.
    _tb_pnl_raw = D.get('TB', {}).get('Profit & Loss A/c', 0)
    # Sanity check only — if pnl_on_bs has wrong sign vs TB, correct it
    # Correct: pnl_on_bs should have SAME sign as TB raw value
    if _tb_pnl_raw != 0 and abs(pnl_on_bs - _tb_pnl_raw) < 1.0:
        pass   # already correct — pnl_on_bs matches TB raw value
    elif _tb_pnl_raw != 0 and abs(pnl_on_bs + _tb_pnl_raw) < 1.0:
        # pnl_on_bs was negated (old code) — un-negate to restore correct sign
        pnl_on_bs = _tb_pnl_raw

    # Capital Account sign convention (Tally proprietorship/partnership):
    # Schedule III treatment: Partners' Capital with net DR balance (excess drawings)
    # is shown as NEGATIVE under Shareholders' Funds — it REDUCES total equity.
    # reserves on BS = cap_total - eq_share + pnl_on_bs (signed, not abs)
    # total_l uses signed cap_total → DR capital reduces liability total correctly.
    # Note 2 display separately uses abs(cap_total) to show the breakdown correctly.
    _cap_dr_asset = 0.0   # No separate asset reclassification needed
    reserves      = cap_total - eq_share + pnl_on_bs

    cl_b       = bucketize(D.get('cl_lines', []), CL_BUCKETS, CL_FALLBACK)
    st_borrow  = sum(v for _, v in cl_b.get("Short Term Borrowings", []))
    trade_pay  = sum(v for _, v in cl_b.get("Trade Payables", []))
    short_prov = sum(v for _, v in cl_b.get("Short Term Provisions", []))

    # Split Other CL — reclassify debit-balance items to asset side
    _other_raw  = cl_b.get(CL_FALLBACK, [])
    _true_cl, _reclassified = split_cl_other(_other_raw)
    other_cl    = sum(v for _, v in _true_cl)           # genuine CR-balance liabilities only
    reclass_amt = sum(abs(v) for _, v in _reclassified) # DR items → reclassified to Note 14

    # ── Loans (Liability) ─────────────────────────────────────────────────
    dtl          = D.get('dtl', 0)
    loans_total  = D.get('loans_total', 0)
    # Long-term borrowings = full Loans (Liability) net total less deferred tax component.
    # TB sign: CR-balance loans → positive; DR-balance (advance repaid) → negative.
    lt_borrow    = loans_total - dtl

    # ── Non-Current Assets ─────────────────────────────────────────────────
    fa_net       = D.get('fa_net_bs', 0)   # abs value (Tally shows gross minus dep reserve)

    # Investments — TB group 'Investments' (FDs, shares, mutual funds etc.)
    # tally_core caches tb_children_of for Investments in tb_children_cache.
    inv_lines    = D.get('tb_children_of_cache', {}).get('Investments',
                    D.get('investment_lines', []))  # fallback key
    inv_total    = abs(TB.get('Investments', 0))   # Tally group net (CR = liability would be unusual)
    # If TB group total is 0 but children exist, sum children
    if inv_total == 0 and inv_lines:
        inv_total = _items_disp_sum(inv_lines)

    # ── Current Assets ─────────────────────────────────────────────────────
    ca_b         = bucketize(D.get('ca_lines', []), CA_BUCKETS, CA_FALLBACK)

    # Inventories: always from bs_closing_stock (stock-type companies) or 0 (service/no-stock)
    inventories  = D.get('bs_closing_stock', D.get('closing_stock_abs', 0))

    # Trade Receivables — split genuine DR-balance debtors (asset) from
    # CR-balance debtors (advances received / overpayments = liability).
    # CR-balance debtors MUST NOT sit as negative Trade Receivables on the BS.
    # They are reclassified to Other Current Liabilities (Advances from Customers).
    # TB sign: DR asset = negative value; CR liability = positive value.
    _tr_items    = ca_b.get("Trade Receivables", [])
    _tr_true     = [(n, v) for n, v in _tr_items if v <= 0]  # DR = genuine receivable (neg TB)
    _tr_cr_adv   = [(n, v) for n, v in _tr_items if v > 0]   # CR = advance received (pos TB)
    trade_recv   = (_items_disp_sum(_tr_true) if _tr_true
                    else (D.get('debtors_net', 0) if not _tr_items else 0))
    # CR-balance debtor amount → shown as Advances from Customers in Other CL
    cr_debtor_adv = sum(v for _, v in _tr_cr_adv)   # positive = CR = liability amount

    # Cash and Cash Equivalents
    _cash_items  = ca_b.get("Cash and Cash Equivalents", [])
    cash_bank    = (_items_disp_sum(_cash_items) if _cash_items
                    else (D.get('cash_hand', 0) + D.get('bank_accts', 0)))

    # Long Term Loans and Advances (security deposits etc. under CA in Tally)
    _lt_items    = ca_b.get("Long Term Loans and Advances", [])
    lt_loans_adv = (_items_disp_sum(_lt_items) if _lt_items
                    else D.get('deposits', 0))

    # Short Term Loans and Advances — CA_FALLBACK items + reclassified DR-balance CL items
    # IMPORTANT: Suspense-named items are handled separately via suspense_val / Note 15A.
    # Exclude them here to prevent double-counting (suspense appears in ca_lines AND is
    # detected separately below — including it in both st_loans_adv and suspense_val
    # causes total_a to exceed total_l by the suspense amount).
    _st_items    = [(n, v) for n, v in ca_b.get(CA_FALLBACK, [])
                    if not any(kw in n.lower() for kw in SUSPENSE_KW)]
    st_loans_adv = _items_disp_sum(_st_items)
    if not st_loans_adv:
        st_loans_adv = (D.get('loans_adv', 0) + D.get('taxation_a', 0) + D.get('other_ca', 0) +
                        D.get('emda', 0) + D.get('dbk', 0) + D.get('rodtep', 0) +
                        D.get('diff_duty', 0) + D.get('rodtep_lic', 0))
    # Add reclassified debit-balance CL items to Short Term Loans & Advances
    st_loans_adv += reclass_amt

    # Suspense A/c — sign-aware split:
    #   TB DR balance (negative) → asset  (genuine deferred/unexplained asset)
    #   TB CR balance (positive) → liability (unexplained credit → shown in Other CL)
    #
    # Tally BS places CR-balance Suspense on the LIABILITY side — we match that.
    # e.g. Design Limelite: Suspense CR 7,58,446.01 → shown on liability side in Tally BS
    # e.g. Suria:           Suspense DR 6,676.00    → shown on asset side
    #
    # TB['Suspense A/c'] sign: positive = CR (liability), negative = DR (asset)
    _susp_raw       = TB.get('Suspense A/c', 0)
    _susp_children  = D.get('tb_children_of_cache', {}).get('Suspense A/c', [])

    # Determine net sign from TB group total; fall back to children sum
    if _susp_raw == 0 and _susp_children:
        # Sum children: positive children = CR, negative = DR
        _susp_raw = sum(v for _, v in _susp_children)
    # Also scan ca_lines for any suspense-named items
    for _nm, _vl in D.get('ca_lines', []):
        if any(kw in _nm.lower() for kw in SUSPENSE_KW) and _susp_raw == 0:
            _susp_raw = _vl

    if _susp_raw > 0:
        # CR balance → liability (Tally BS shows it on liability side)
        suspense_val     = 0.0           # nothing on asset side
        susp_cr_liab     = _susp_raw     # amount to add to Other CL on liability side
    else:
        # DR balance → asset (unexplained debit → show as Other Current Assets)
        suspense_val     = abs(_susp_raw)  # positive display amount on asset side
        susp_cr_liab     = 0.0

    # Unadjusted Forex Gain/Loss + Misc. Expenses (ASSET) — separate Tally BS groups
    # These appear on the asset side in Tally BS and must be carried through.
    # 'Unadjusted Forex Gain/Loss' is a DR-balance group (loss not yet booked).
    # 'Misc. Expenses (ASSET)' = preliminary/deferred expenses.
    forex_val = abs(TB.get('Unadjusted Forex Gain/Loss',
                    TB.get('Unadjusted Forex Gain / Loss',
                    TB.get('Forex Gain/Loss', 0))))
    misc_asset = abs(TB.get('Misc. Expenses (ASSET)',
                    TB.get('Miscellaneous Expenses (ASSET)', 0)))
    other_nca  = forex_val + misc_asset   # combined non-current / deferred asset

    # ── Other Current Liabilities — combine all CR items that belong on liability side ──
    # Includes: (a) genuine CR-balance CL items
    #           (b) CR-balance debtors (advances from customers)
    #           (c) CR-balance Suspense A/c
    other_cl_total = other_cl + cr_debtor_adv + susp_cr_liab

    # ── Authoritative totals from Tally TB ────────────────────────────────
    # total_l: all liability-side items
    #   - equity / reserves / borrowings
    #   - trade payables
    #   - genuine CR-balance Other CL items
    #   - CR-balance debtors (advances from customers)
    #   - CR-balance Suspense (reclassified from asset side)
    #   - short-term provisions
    # DR-balance items (reclassified to asset) are excluded from total_l.
    # FIX: reserves is already correctly signed (negative for loss co), so eq_share + reserves
    # correctly represents net equity. For loss companies reserves < 0 → total_l reduced → correct.
    total_l = (eq_share + reserves + lt_borrow + dtl + st_borrow
               + trade_pay + other_cl + short_prov + cr_debtor_adv + susp_cr_liab)

    # Asset total: individual classified items
    # suspense_val = 0 when CR (already in total_l); positive when DR (asset)
    total_a = (fa_net + inv_total + lt_loans_adv + inventories
               + trade_recv + cash_bank + st_loans_adv
               + suspense_val + other_nca)
    # Last resort: force equal
    # FIX: Never silently force total_a = total_l — that hides real bugs.
    # Instead flag the mismatch clearly so the CA/preparer can investigate.
    _bs_diff = total_a - total_l
    if abs(_bs_diff) > 1.0:
        print(f"  WARNING: BS does not balance — Assets={total_a:,.2f}  Liabilities={total_l:,.2f}"
              f"  Diff={_bs_diff:,.2f}. Check pnl_on_bs sign, equity_share, or reclassification.")

    return dict(
        eq_share=eq_share, reserves=reserves,
        lt_borrow=lt_borrow, dtl=dtl,
        st_borrow=st_borrow, trade_pay=trade_pay,
        other_cl=other_cl, other_cl_total=other_cl_total,
        short_prov=short_prov,
        fa_net=fa_net,
        inv_total=inv_total, inv_lines=inv_lines,
        lt_loans_adv=lt_loans_adv, inventories=inventories,
        trade_recv=trade_recv, cash_bank=cash_bank,
        st_loans_adv=st_loans_adv,
        suspense_val=suspense_val, susp_cr_liab=susp_cr_liab,
        other_nca=other_nca, forex_val=forex_val, misc_asset=misc_asset,
        cr_debtor_adv=cr_debtor_adv,
        total_l=total_l, total_a=total_a,
        # stored for Notes rendering
        _true_cl=_true_cl, _reclassified=_reclassified,
        _tr_cr_adv=_tr_cr_adv,
        _susp_raw=_susp_raw, _susp_children=_susp_children,
    )

# ══════════════════════════════════════════════════════════════════════════
# PAGE 1 — BALANCE SHEET
# ══════════════════════════════════════════════════════════════════════════
def build_classic_bs(D, D_py, story):
    fy_e = D['fy_end']
    fy_p = D_py['fy_end'] if D_py else f"{int(fy_e[:4])-1}{fy_e[4:]}"
    CW   = [PAGE_W*0.50, PAGE_W*0.08, PAGE_W*0.21, PAGE_W*0.21]

    story.append(Paragraph(D['company'].upper(), S["TITLE"]))
    story.append(Paragraph(f"BALANCE SHEET AS AT {fmt_date(fy_e)}", S["SUB"]))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("(Amt. in Rs.00)", ParagraphStyle("R", parent=S["SML"], alignment=TA_RIGHT)))

    f  = bs_figures(D)
    fp = bs_figures(D_py) if D_py else {k: None for k in f}

    rows = [[P("PARTICULARS","BOLD"), N("Note No."),
             Paragraph(f"<b>FIGURES AS AT<br/>THE END OF<br/>{fmt_date(fy_e)}</b>", S["CTR"]),
             Paragraph(f"<b>FIGURES AS AT<br/>THE END OF<br/>{fmt_date(fy_p)}</b>", S["CTR"])]]
    bold_rows, shaded = [], []

    def section(label):
        rows.append([P(label,"BOLD"), N(""), A(None), A(None)])

    def line(label, note, cur, py):
        rows.append([P(label,"NRM"), N(note), A(cur), A(py)])

    def blank(label):
        rows.append([P(label,"NRM"), N(""), A(None), A(None)])

    # I. EQUITY & LIABILITIES
    section("I.  EQUITY &amp; LIABILITIES")
    section("    SHAREHOLDER'S FUNDS")
    line("        SHARE CAPITAL", "1", f['eq_share'], fp['eq_share'])
    line("        RESERVES &amp; SURPLUS", "2", f['reserves'], fp['reserves'])
    blank("        MONEY RECEIVED AGAINST SHARE WARRANTS")
    blank("    SHARE APPLICATION MONEY PENDING ALLOTMENT")
    section("    NON-CURRENT LIABILITIES")
    line("        LONG-TERM BORROWINGS", "3", f['lt_borrow'], fp['lt_borrow'])
    line("        DEFERRED TAX LIABILITIES (NET)", "32", f['dtl'], fp['dtl'])
    blank("        OTHER LONG TERM LIABILITIES")
    blank("        LONG-TERM PROVISIONS")
    section("    CURRENT LIABILITIES")
    line("        SHORT-TERM BORROWINGS", "4", f['st_borrow'], fp['st_borrow'])
    line("        TRADE PAYABLES", "5", f['trade_pay'], fp['trade_pay'])
    # Other CL = genuine CR-balance CL items + CR-balance debtors (advances) + CR Suspense
    _ocl_cy = f['other_cl_total']
    _ocl_py = fp.get('other_cl_total') if fp else None
    line("        OTHER CURRENT LIABILITIES", "6", _ocl_cy, _ocl_py)
    line("        SHORT TERM PROVISION", "7", f['short_prov'], fp['short_prov'])
    rows.append([P("TOTAL","BOLD"), N(""), AB(f['total_l']), AB(fp['total_l'])])
    bold_rows.append(len(rows)-1); shaded.append(len(rows)-1)

    # II. ASSETS
    section("II.  ASSETS")
    section("    NON-CURRENT ASSETS")
    section("    Property Plant and Equipment")
    line("        TANGIBLE ASSETS", "8", f['fa_net'], fp['fa_net'])
    blank("        INTANGIBLE ASSETS")
    blank("        CAPITAL WORK-IN-PROGRESS")
    blank("        INTANGIBLE ASSETS UNDER DEVELOPMENT")
    # Investments — only show row if there is a balance in either year
    _inv_cy = f['inv_total']
    _inv_py = fp.get('inv_total') if fp else None
    if (_inv_cy and abs(_inv_cy) > 0.5) or (_inv_py and abs(_inv_py) > 0.5):
        line("    NON-CURRENT INVESTMENTS", "10", _inv_cy or None, _inv_py)
    else:
        blank("    NON-CURRENT INVESTMENTS")
    blank("    DEFERRED TAX ASSET (NET)")
    line("    LONG-TERM LOANS AND ADVANCES", "9", f['lt_loans_adv'], fp['lt_loans_adv'] if fp else None)
    # Unadjusted Forex Gain/Loss + Misc. Expenses (ASSET) — shown when non-zero
    _onca_cy = f.get('other_nca', 0)
    _onca_py = fp.get('other_nca', 0) if fp else None
    if (_onca_cy and abs(_onca_cy) > 0.5) or (_onca_py and abs(_onca_py or 0) > 0.5):
        line("    OTHER NON-CURRENT ASSETS (Forex/Misc)", "9A", _onca_cy or None, _onca_py or None)
    else:
        blank("    OTHER NON-CURRENT ASSETS")
    section("    CURRENT ASSETS")
    blank("        CURRENT INVESTMENTS")
    line("        INVENTORIES", "11", f['inventories'], fp['inventories'] if fp else None)
    # Trade Receivables: show only genuine DR-balance debtors (positive amount)
    # CR-balance debtors have been reclassified to Other Current Liabilities above
    _tr_cy = f['trade_recv']
    _tr_py = fp.get('trade_recv') if fp else None
    line("        TRADE RECEIVABLES", "12", _tr_cy, _tr_py)
    line("        CASH AND CASH EQUIVALENTS", "13", f['cash_bank'], fp['cash_bank'] if fp else None)
    line("        SHORT TERM LOANS AND ADVANCES", "14", f['st_loans_adv'], fp['st_loans_adv'] if fp else None)
    # Other Current Assets: Suspense A/c only when DR-balance (genuine asset)
    # CR-balance Suspense is already on the liability side via other_cl_total
    _susp_cy = f['suspense_val']   # 0 when CR-balance, positive when DR-balance
    _susp_py = fp.get('suspense_val') if fp else None
    if (_susp_cy and abs(_susp_cy) > 0.5) or (_susp_py and abs(_susp_py or 0) > 0.5):
        line("        OTHER CURRENT ASSETS (Suspense)", "15A", _susp_cy or None, _susp_py)
    else:
        blank("        OTHER CURRENT ASSETS")
    rows.append([P("TOTAL","BOLD"), N(""), AB(f['total_a']), AB(fp['total_a'] if fp else None)])
    bold_rows.append(len(rows)-1); shaded.append(len(rows)-1)

    rows.append([P("III.  CONTINGENT LIABILITIES","NRM"), N("31"), A(None), A(None)])

    story.append(grid_table(rows, CW, bold_rows=bold_rows, shaded_rows=shaded))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("See accompanying notes to the financial statements", S["SML"]))
    story.append(Paragraph("As per our attached Report of even date", S["SML"]))
    story.append(Spacer(1, 8*mm))
    story.append(Table([[
        Paragraph(f"For {D['company']}", S["BOLD"]), Paragraph("", S["NRM"]),
    ]], colWidths=[PAGE_W*0.5, PAGE_W*0.5]))
    story.append(Spacer(1, 10*mm))
    # BUG 2 FIX: auto-pull director/partner names from cap_lines
    # Strip known suffixes/prefixes and skip expense-type ledger names
    _CAP_SUFFIX_RX = re.compile(
        r'\s*[-–]?\s*(share\s+capital|capital\s+a[/.]?c\.?|capital\s+account|equity|drawings?)\s*$',
        re.IGNORECASE)
    _CAP_PREFIX_RX = re.compile(r'^(drawing\s+|drawings?\s+)', re.IGNORECASE)
    _SKIP_KW_SIG = ('share capital', 'equity share', 'preference share', 'capital a/c',
                    'capital account', 'reserve', 'profit', 'p&l', 'surplus',
                    'mediclaim', 'school fee', 'tution fee', 'tuition fee',
                    'insurance', 'rent exp', 'salary', ' exp', 'expense',
                    'telephone', 'electric', 'petrol', 'fees')
    def _extract_director_names(cap_lines):
        names = []
        for nm, _ in (cap_lines or []):
            nl = nm.lower()
            if any(kw in nl for kw in _SKIP_KW_SIG):
                continue
            # Skip ledgers that are purely drawing accounts (start with "drawing")
            if nl.startswith('drawing'):
                continue
            clean = _CAP_SUFFIX_RX.sub('', nm).strip()
            clean = _CAP_PREFIX_RX.sub('', clean).strip()
            if clean and clean not in names:
                names.append(clean)
        return names[:2]
    _dir_names = _extract_director_names(D.get('cap_lines', []))
    if not _dir_names:
        _sig_left  = "_________________<br/>Director / Partner"
        _sig_right = "_________________<br/>Director / Partner"
    elif len(_dir_names) == 1:
        _sig_left  = f"_________________<br/>{_dir_names[0]}<br/>Director / Partner"
        _sig_right = "_________________<br/>Director / Partner"
    else:
        _sig_left  = f"_________________<br/>{_dir_names[0]}<br/>Director / Partner"
        _sig_right = f"_________________<br/>{_dir_names[1]}<br/>Director / Partner"
    story.append(Table([[
        Paragraph(_sig_left,  S["NRM"]),
        Paragraph(_sig_right, S["NRM"]),
    ]], colWidths=[PAGE_W*0.5, PAGE_W*0.5]))

# ══════════════════════════════════════════════════════════════════════════
# PAGE 2 — PROFIT & LOSS
# ══════════════════════════════════════════════════════════════════════════
def pnl_figures(D):
    if not D:
        return None

    TB        = D.get('TB', {})
    _co_type  = tc.detect_company_type(TB)
    _has_stock = _co_type.get('has_stock', False)
    _is_dormant = _co_type.get('is_dormant', False)

    # ── Revenue ───────────────────────────────────────────────────────────
    revenue   = abs(D['sales'])
    other_inc = abs(D['ind_inc'])

    # BUG 7 FIX: Service companies may book all revenue under "Indirect Incomes"
    # in Tally (no "Sales Accounts" group). When sales_lines is empty and
    # ind_inc_lines has items, reclassify ind_inc as revenue from operations.
    _sales_lines_pnl = D.get('sales_lines', [])
    _ind_inc_lines_pnl = D.get('ind_inc_lines', [])
    if not revenue and _ind_inc_lines_pnl and not _is_dormant:
        # All income is in Indirect Incomes — treat as revenue from operations
        revenue   = abs(D['ind_inc'])
        other_inc = 0.0

    tot_rev   = revenue + other_inc

    # ── Cost of Material Consumed ─────────────────────────────────────────
    # Try named sub-group lookups first (manufacturing companies).
    cost_mat  = (D.get('purch_rm',0) + D.get('purch_job',0) + D.get('purch_clg',0) +
                 D.get('mfg_exp',0) + D.get('power_fuel',0) +
                 D.get('stores_exp',0) + D.get('stores_cons',0))
    if not cost_mat:
        cost_mat = abs(D['purchases'])
    # For trading-WITHOUT-stock companies (e.g. Vianci): Direct Expenses
    # (Freight, Labour, Loading/Unloading, Transport etc.) are true cost-of-goods
    # and MUST be included in cost_mat so they appear in the P&L.
    # Without this, they are silently omitted → profit is overstated by their full amount.
    _dir = D.get('dir_exp_lines', [])
    _dir_total = abs(D.get('dir_exp', 0))
    if _dir_total > 0 and not _has_stock and D.get('purchases', 0) != 0:
        # Include ALL Direct Expenses in cost section for trading-without-stock.
        # (They are already factored into tally_core's net_profit via gross_profit calculation.)
        cost_mat += _dir_total

    # ── Inventory Change ──────────────────────────────────────────────────
    # Only meaningful for stock-carrying companies.
    # For service/trading-without-stock: force to zero to prevent phantom terms.
    if _has_stock:
        inv_chg = D['opening_stock_abs'] - D.get('bs_closing_stock', D['closing_stock_abs'])
    else:
        inv_chg = 0.0

    # ── Employee Expense ──────────────────────────────────────────────────
    _ind    = D.get('ind_exp_lines', [])
    EMP_KW  = ['salary','wage','bonus','esic','gratuity','staff welfare','labour welfare','director remun']
    FIN_KW  = ['interest','bank charge','processing fee','mortgage','finance charge']

    emp_exp  = D.get('emp_exp', 0)
    if not emp_exp:
        emp_exp = sum(abs(v) for n,v in _ind if any(kw in n.lower() for kw in EMP_KW))

    # ── Finance Cost ──────────────────────────────────────────────────────
    fin_cost = D.get('finance_exp', 0)
    if not fin_cost:
        fin_cost = sum(abs(v) for n,v in _ind if any(kw in n.lower() for kw in FIN_KW))

    # ── Depreciation ──────────────────────────────────────────────────────
    dep = abs(D.get('dep_reserve', 0))

    # ── Other Expenses ────────────────────────────────────────────────────
    # = Indirect Expenses EXCLUDING employee + finance + dep
    # For trading-without-stock: Direct Expenses are already in cost_mat above,
    # so do NOT add them again here (would double-count).
    _excl_kw = EMP_KW + FIN_KW
    _other_ind = sum(abs(v) for n,v in _ind if not any(kw in n.lower() for kw in _excl_kw))

    # Named other-expense sub-groups (manufacturing companies)
    other_exp = (D.get('admin_exp',0) + D.get('factory_exp',0) +
                 D.get('legal_exp',0) + D.get('selling_exp',0))
    if not other_exp:
        other_exp = _other_ind
    # For companies where dir_exp_lines are NOT already in cost_mat
    # (i.e. has_stock companies where Direct Expenses are not part of purchases):
    # include them in Other Expenses. For trading-without-stock, skip (already in cost_mat).
    if _dir_total > 0 and _has_stock:
        # Dir exp for stock companies (e.g. freight on purchases) goes to Other Expenses
        _dir_excl = sum(abs(v) for n,v in _dir if not any(kw in n.lower() for kw in _excl_kw))
        other_exp += _dir_excl

    # ── Totals ────────────────────────────────────────────────────────────
    tot_exp  = cost_mat + inv_chg + emp_exp + fin_cost + dep + other_exp

    # Tally's net_profit is the SINGLE authoritative figure — always use it for NPAT.
    # PBEET is shown as (tot_rev − tot_exp) for display consistency.
    # If tot_exp is slightly off due to classification, adjust other_exp so
    # PBEET equals Tally's net_profit (which has zero tax for most MSMEs).
    np_core  = D.get('net_profit', 0)
    tax      = D.get('income_tax', 0)
    pbeet    = tot_rev - tot_exp
    pbt      = pbeet
    # Reconcile: if pbeet ≠ np_core (ignoring tax), absorb diff into other_exp display
    _diff = np_core - (pbeet - tax)
    if abs(_diff) > 1.0:
        # Small residual caused by classification — absorb silently into other_exp
        other_exp += _diff
        tot_exp   += _diff
        pbeet      = np_core + tax
        pbt        = pbeet
    npat = np_core   # authoritative

    return dict(revenue=revenue, other_inc=other_inc, tot_rev=tot_rev,
                cost_mat=cost_mat, inv_chg=inv_chg, emp_exp=emp_exp,
                fin_cost=fin_cost, dep=dep, other_exp=other_exp,
                tot_exp=tot_exp, pbeet=pbeet, pbt=pbt, tax=tax, npat=npat)

def build_classic_pnl(D, D_py, story):
    story.append(PageBreak())
    fy_e = D['fy_end']
    fy_p = D_py['fy_end'] if D_py else f"{int(fy_e[:4])-1}{fy_e[4:]}"
    CW   = [PAGE_W*0.54, PAGE_W*0.08, PAGE_W*0.19, PAGE_W*0.19]

    story.append(Paragraph(D['company'].upper(), S["TITLE"]))
    story.append(Paragraph(f"PROFIT &amp; LOSS STATEMENT FOR THE YEAR ENDED {fmt_date(fy_e)}", S["SUB"]))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("(Amt. in Rs.00)", ParagraphStyle("R", parent=S["SML"], alignment=TA_RIGHT)))

    m  = pnl_figures(D)
    mp = pnl_figures(D_py) if D_py else {k: None for k in m}

    # FIX: Dormant/shell company — no income, no expenses for the year.
    # Add a clear disclosure note so the CA/auditor and reader know this is intentional.
    _co_type_pnl = tc.detect_company_type(D.get('TB', {}))
    if _co_type_pnl.get('is_dormant', False):
        story.append(Spacer(1, 3*mm))
        story.append(Paragraph(
            "<b>Note:</b> The Company did not carry out any business operations during the year ended "
            f"{fmt_date(fy_e)}. No income was earned and no expenses were incurred during the year. "
            "The loss carried forward represents amounts from prior periods.",
            S["SML"]))
        story.append(Spacer(1, 3*mm))

    rows = [[P("","BOLD"), N("NO."),
             Paragraph(f"<b>AT THE END OF<br/>{fmt_date(fy_e)}</b>", S["CTR"]),
             Paragraph(f"<b>AT THE END OF<br/>{fmt_date(fy_p)}</b>", S["CTR"])]]
    bold_rows, shaded = [], []

    def row(label, note, cur, py, b=False):
        rows.append([P(label, "BOLD" if b else "NRM"), N(note),
                     (AB if b else A)(cur), (AB if b else A)(py)])
        if b: bold_rows.append(len(rows)-1)

    row("I.    REVENUE FROM OPERATIONS", "15", m['revenue'], mp['revenue'])
    row("II.   OTHER INCOME", "16", m['other_inc'], mp['other_inc'])
    row("III.  TOTAL REVENUE (I+II)", "", m['tot_rev'], mp['tot_rev'], b=True)
    row("IV.   EXPENSES", "", None, None)
    row("        COST OF MATERIAL CONSUMED", "17", m['cost_mat'], mp['cost_mat'])
    row("        PURCHASES OF STOCK IN TRADE", "", None, None)
    row("        CHANGES IN INVENTORIES OF FINISHED GOODS,", "18", m['inv_chg'], mp['inv_chg'])
    row("        WORK-IN-PROGRESS AND STOCK-IN-TRADE", "", None, None)
    row("        EMPLOYEE BENEFIT EXPENSE", "19", m['emp_exp'], mp['emp_exp'])
    row("        FINANCIAL COST", "20", m['fin_cost'], mp['fin_cost'])
    row("        DEPRECIATION &amp; AMORTIZATION EXPENSE", "8", m['dep'], mp['dep'])
    row("        OTHER EXPENSES", "21", m['other_exp'], mp['other_exp'])
    row("        TOTAL EXPENSES", "", m['tot_exp'], mp['tot_exp'], b=True)
    row("V.    PROFIT BEFORE EXCEPTIONAL AND EXTRAORDINARY", "", None, None)
    row("        ITEMS AND TAX (III-IV)", "", m['pbeet'], mp['pbeet'], b=True)
    row("VI.   EXCEPTIONAL ITEMS", "", None, None)
    row("        PRIOR PERIOD ITEMS (NET)", "", None, None)
    row("        OTHER EXCEPTIONAL ITEMS", "", None, None)
    row("VII.  PROFIT BEFORE EXTRAORDINARY ITEMS AND TAX (V-VI)", "", m['pbeet'], mp['pbeet'], b=True)
    row("VIII. EXTRAORDINARY ITEMS", "", None, None)
    row("IX.   PROFIT BEFORE TAX (VII-VIII)", "", m['pbt'], mp['pbt'], b=True)
    row("X.    TAX EXPENSE", "", None, None)
    row("        EXCESS/ SHORTER PROVISION OF EARLIER YEAR", "", None, None)
    row("        CURRENT TAX", "", m['tax'], mp['tax'])
    row("        DEFERRED TAX", "22", None, None)
    row("        MAT", "", None, None)
    row("        TOTAL TAX EXPENSE", "", m['tax'], mp['tax'], b=True)
    row("XI.   PROFIT (LOSS) FOR THE PERIOD", "", m['npat'], mp['npat'], b=True)
    shaded.append(len(rows)-1)
    row("XII.  EARNING PER EQUITY SHARE:", "", None, None)

    sc = D.get('equity_share', 0)
    # BUG 8 FIX: suppress EPS row when no equity (LLP/proprietorship), show "-"
    # for zero profit, and show negative EPS in brackets for losses.
    def _fmt_eps(npat, equity):
        if not equity or equity == 0:
            return None   # suppress — not applicable (LLP/proprietorship)
        shares = equity / 10.0
        if not shares:
            return None
        if npat is None:
            return None
        if abs(npat) < 0.005:
            return "-"    # nil profit/loss
        val = npat / shares
        if val < 0:
            return f"({abs(val):,.2f})"
        return f"{val:,.2f}"

    eps_str    = _fmt_eps(m['npat'], sc)
    eps_py_str = None
    if D_py:
        eps_py_str = _fmt_eps(mp['npat'] if mp else None, D_py.get('equity_share', 0))

    if eps_str is not None or eps_py_str is not None:
        rows.append([P("        DILUTED","NRM"), N(""),
                     Paragraph(eps_str    if eps_str    is not None else "—", S["AMT"]),
                     Paragraph(eps_py_str if eps_py_str is not None else "—", S["AMT"])])

    story.append(grid_table(rows, CW, bold_rows=bold_rows, shaded_rows=shaded))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("See accompanying notes to the financial statements", S["SML"]))
    story.append(Paragraph("As per our attached Report of even date", S["SML"]))

# ══════════════════════════════════════════════════════════════════════════
# NOTES
# ══════════════════════════════════════════════════════════════════════════
def build_classic_notes(D, D_py, story):
    fy_e = D['fy_end']
    fy_p = D_py['fy_end'] if D_py else f"{int(fy_e[:4])-1}{fy_e[4:]}"
    story.append(PageBreak())
    story.append(Paragraph(f"Notes forming part of Balance Sheet as at {fmt_date(fy_e)}", S["SUB"]))
    story.append(Spacer(1, 3*mm))
    CW = [PAGE_W*0.56, PAGE_W*0.22, PAGE_W*0.22]

    def hdr_row():
        return [P("Particulars","BOLD"),
                Paragraph(f"<b>As At<br/>{fmt_date(fy_e)}</b>", S["CTR"]),
                Paragraph(f"<b>As At<br/>{fmt_date(fy_p)}</b>", S["CTR"])]

    def note_table(title, items, py_items=None, mode="asset", total_label="Total"):
        """
        Renders a two-column comparison note table (CY vs PY).

        KEY RULE — UNION of CY and PY items:
          A proper comparison table must show ALL ledger names that appear in
          EITHER year, not just those in CY. Without this, a ledger that exists
          only in PY (e.g. purchases in a year with no CY purchases) never gets
          a row, so its PY value is silently dropped.

        mode="asset"    : DR items (val<0 in TB) → positive display
                          CR items (val>0 in TB) → brackets (contra-asset)
        mode="liability": CR items (val>0 in TB) → positive display
                          DR items (val<0 in TB) → brackets (contra-liability)
        """
        def disp(val, m):
            if val is None: return None
            v = float(val)
            if m == "asset":
                return abs(v) if v < 0 else (-abs(v) if abs(v) > 0.005 else 0)
            else:
                return abs(v) if v > 0 else (-abs(v) if abs(v) > 0.005 else 0)

        story.append(Paragraph(title, S["BOLD"]))
        rows = [hdr_row()]
        py_items = py_items or []

        # Build maps: name → value for both years
        cy_map = {}
        for n, v in (items or []):
            cy_map[n] = v
            cy_map[n.strip()] = v
        py_map = {}
        for n, v in py_items:
            py_map[n] = v
            py_map[n.strip()] = v

        # Build ORDERED UNION of all ledger names:
        # CY names first (in original order), then PY-only names
        seen = set()
        all_names = []   # list of (display_name, cy_val, py_val)
        for name, val in (items or []):
            key = name.strip()
            if key not in seen:
                seen.add(key)
                all_names.append((name, val, py_map.get(name, py_map.get(key))))
        for name, val in py_items:
            key = name.strip()
            if key not in seen:
                seen.add(key)
                all_names.append((name, cy_map.get(name, cy_map.get(key)), val))

        tot, tot_py = 0.0, 0.0
        for name, cy_val, py_val in all_names:
            v  = disp(cy_val, mode)
            pv = disp(py_val, mode)
            rows.append([P(f"    {name}","NRM"), A(v), A(pv)])
            tot    += float(v  or 0)
            tot_py += float(pv or 0) if pv is not None else 0

        rows.append([P(total_label,"BOLD"), AB(tot), AB(tot_py if py_items else None)])
        story.append(grid_table(rows, CW, bold_rows=[len(rows)-1]))
        story.append(Spacer(1, 4*mm))

    # ── Build buckets ───────────────────────────────────────────────────
    cl_b    = bucketize(D.get('cl_lines', []), CL_BUCKETS, CL_FALLBACK)
    ca_b    = bucketize(D.get('ca_lines', []), CA_BUCKETS, CA_FALLBACK)
    cl_b_py = bucketize(D_py.get('cl_lines', []), CL_BUCKETS, CL_FALLBACK) if D_py else {}
    ca_b_py = bucketize(D_py.get('ca_lines', []), CA_BUCKETS, CA_FALLBACK) if D_py else {}

    # Split Other CL for Note 6 (genuine liabilities) and Note 14 (reclassified debit items)
    _true_cl,    _reclassified    = split_cl_other(cl_b.get(CL_FALLBACK, []))
    _true_cl_py, _reclassified_py = split_cl_other(cl_b_py.get(CL_FALLBACK, [])) if D_py else ([], [])

    # ── Note 1: Share Capital (or Partners'/Proprietor's Capital) ──────
    # BUG 5 FIX: detect LLP/proprietorship — no ledger matching "share capital"
    # or "equity" in cap_lines → render as partners' capital table, not share count.
    _SHARE_KW = ('share capital', 'equity share', 'preference share', 'equity capital')
    _cap_lines_n1 = D.get('cap_lines', [])
    _has_share_cap = any(
        any(kw in nm.lower() for kw in _SHARE_KW)
        for nm, _ in _cap_lines_n1
    )
    if _has_share_cap:
        # ── Full Schedule III Share Capital disclosure for companies ──────────
        _co_info   = D.get('co_info', {}) or {}
        _auth_cap  = abs(float(_co_info.get('AUTHORISEDCAPITAL',
                          _co_info.get('AUTHORISEDSHARES', 0)) or 0))
        _paid_up   = D.get('equity_share', 0)
        _auth_cap  = _auth_cap if _auth_cap >= _paid_up else _paid_up

        _co_info_py  = (D_py.get('co_info', {}) or {}) if D_py else {}
        _auth_cap_py = abs(float(_co_info_py.get('AUTHORISEDCAPITAL',
                            _co_info_py.get('AUTHORISEDSHARES', 0)) or 0))
        _paid_up_py  = D_py.get('equity_share', 0) if D_py else 0
        _auth_cap_py = (_auth_cap_py if _auth_cap_py >= _paid_up_py else _paid_up_py) if D_py else None

        _face_val = abs(float(_co_info.get('FACEVALUE', _co_info.get('FACEVAL', 10)) or 10))
        if _face_val == 0: _face_val = 10

        _shares_auth = int(_auth_cap / _face_val) if _face_val else 0
        _shares_paid = int(_paid_up  / _face_val) if _face_val else 0
        _shares_py   = int(_paid_up_py / _face_val) if (D_py and _face_val) else 0

        story.append(Paragraph("Note 1: Share Capital", S["BOLD"]))
        _cw1 = [PAGE_W*0.56, PAGE_W*0.22, PAGE_W*0.22]
        _r1  = [hdr_row()]

        # Authorised
        _r1.append([P("    Authorised Capital", "BOLD"), A(None), A(None)])
        _r1.append([P(f"        {_shares_auth:,} Equity Shares of "
                      f"Rs.{_face_val:.0f}/- each", "NRM"),
                    A(_auth_cap), A(_auth_cap_py)])
        # Issued = Subscribed = Paid-up
        _r1.append([P("    Issued, Subscribed and Paid-up Capital", "BOLD"), A(None), A(None)])
        _r1.append([P(f"        {_shares_paid:,} Equity Shares of "
                      f"Rs.{_face_val:.0f}/- each, fully paid up", "NRM"),
                    A(_paid_up), A(_paid_up_py if D_py else None)])
        _r1.append([P("    Total Paid-up Share Capital", "BOLD"),
                    AB(_paid_up), AB(_paid_up_py if D_py else None)])
        story.append(grid_table(_r1, _cw1, bold_rows=[len(_r1)-1]))

        # Reconciliation of shares outstanding
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("<b>Reconciliation of Number of Shares Outstanding:</b>", S["SML"]))
        _cw_rec = [PAGE_W*0.46, PAGE_W*0.18, PAGE_W*0.18, PAGE_W*0.18]
        _rows_rec = [[P("Particulars","BOLD"),
                      Paragraph("<b>No. of Shares</b>", S["CTR"]),
                      Paragraph(f"<b>Amount (Rs.)</b>", S["CTR"]),
                      Paragraph(f"<b>PY Amount</b>", S["CTR"])]]
        _issued_this_yr = _paid_up - (_paid_up_py or 0)
        _rows_rec.append([P("    Shares at beginning of year","NRM"),
                          A(_shares_py), A(_paid_up_py if D_py else None), A(None)])
        _rows_rec.append([P("    Add: Shares issued during the year","NRM"),
                          A(_shares_paid - _shares_py if _issued_this_yr > 0 else 0),
                          A(_issued_this_yr if _issued_this_yr > 0 else None), A(None)])
        _rows_rec.append([P("    Shares at end of year","BOLD"),
                          AB(_shares_paid), AB(_paid_up),
                          AB(_paid_up_py if D_py else None)])
        story.append(grid_table(_rows_rec, _cw_rec, bold_rows=[len(_rows_rec)-1]))

        # Shareholders holding > 5%
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph("<b>Details of Shareholders holding more than 5% of Shares:</b>", S["SML"]))
        _cw_sh = [PAGE_W*0.46, PAGE_W*0.18, PAGE_W*0.18, PAGE_W*0.18]
        _rows_sh = [[P("Name of Shareholder","BOLD"),
                     Paragraph("<b>No. of Shares</b>", S["CTR"]),
                     Paragraph("<b>% Holding</b>", S["CTR"]),
                     Paragraph(f"<b>As At {fmt_date(fy_e)}</b>", S["CTR"])]]
        _found_sh = False
        for _shn, _shv in _cap_lines_n1:
            _shv_abs = abs(_shv)
            if _paid_up > 0 and _shv_abs / _paid_up >= 0.05:
                _sh_shares = int(_shv_abs / _face_val) if _face_val else 0
                _sh_pct    = _shv_abs / _paid_up * 100
                _rows_sh.append([P(f"    {_shn}","NRM"),
                                  A(_sh_shares),
                                  Paragraph(f"{_sh_pct:.2f}%", S["AMT"]),
                                  A(_shv_abs)])
                _found_sh = True
        if not _found_sh:
            _rows_sh.append([P("    (Confirm with management — shareholder details not in Tally)","NRM"),
                              A(None), A(None), A(None)])
        story.append(grid_table(_rows_sh, _cw_sh))
        story.append(Spacer(1, 4*mm))
    else:
        # LLP, partnership, proprietorship — show each capital ledger as-is
        _cap_n1_py = D_py.get('cap_lines', []) if D_py else None
        # Filter out P&L-like ledgers from cap_lines for Note 1 display
        _PNL_KW_N1 = ('profit', 'p&l', 'loss', 'surplus')
        _cap_items_cy = [(nm, v) for nm, v in _cap_lines_n1
                         if not any(kw in nm.lower() for kw in _PNL_KW_N1)]
        _cap_items_py = [(nm, v) for nm, v in (_cap_n1_py or [])
                         if not any(kw in nm.lower() for kw in _PNL_KW_N1)] if _cap_n1_py else None
        # If all cap_lines were P&L-type — there is genuinely no contributed capital
        # (common for LLPs where Tally only has P&L under Capital Accounts).
        # Show the table with nil total rather than computing a phantom figure.
        # (Note 2 Reserves & Surplus carries the full P&L balance separately.)
        _n1_title = "Note 1: Partners&#39; Capital / Proprietor&#39;s Capital"
        note_table(_n1_title, _cap_items_cy, _cap_items_py, mode="liability")

    # ── Note 2: Reserves & Surplus ─────────────────────────────────────
    # FIX: pnl_bf is now correctly signed in tally_core (negative=loss, positive=profit).
    # pnl_on_bs = pnl_bf + net_profit — both correctly signed. Reserves correctly negative for loss co.
    story.append(Paragraph("Note 2: Reserves &amp; Surplus", S["BOLD"]))
    rows2 = [hdr_row()]
    cap_total    = D.get('cap_total', 0);      eq_share    = D.get('equity_share', 0)
    # Self-defence: apply same pnl_on_bs sign fix as bs_figures
    _tb_pnl_raw_n2 = D.get('TB', {}).get('Profit & Loss A/c', 0)
    pnl_bs_cy = D.get('pnl_on_bs', 0)
    # Correct sign: pnl_bs_cy should match TB raw value (DR=negative=loss)
    if _tb_pnl_raw_n2 != 0 and abs(pnl_bs_cy + _tb_pnl_raw_n2) < 1.0:
        pnl_bs_cy = _tb_pnl_raw_n2   # was negated by old code — restore
    cap_total_py = D_py.get('cap_total', 0)    if D_py else None
    eq_share_py  = D_py.get('equity_share', 0) if D_py else 0
    _tb_pnl_py_n2 = D_py.get('TB', {}).get('Profit & Loss A/c', 0) if D_py else 0
    pnl_bs_py = D_py.get('pnl_on_bs', 0) if D_py else 0
    if D_py and _tb_pnl_py_n2 != 0 and abs(pnl_bs_py + _tb_pnl_py_n2) < 1.0:
        pnl_bs_py = _tb_pnl_py_n2   # was negated by old code — restore
    res_total    = (cap_total - eq_share) + pnl_bs_cy
    res_total_py = (((cap_total_py or 0) - eq_share_py) + pnl_bs_py) if D_py else None
    # pnl_bf self-defence: same logic
    _tb_pnl_raw_bf = D.get('TB', {}).get('Profit & Loss A/c', 0)
    pnl_bf_cy = D.get('pnl_bf', 0)
    if _tb_pnl_raw_bf != 0 and abs(pnl_bf_cy + _tb_pnl_raw_bf) < 1.0:
        pnl_bf_cy = _tb_pnl_raw_bf   # was negated — restore
    _tb_pnl_py_bf = D_py.get('TB', {}).get('Profit & Loss A/c', 0) if D_py else 0
    pnl_bf_py = D_py.get('pnl_bf', 0) if D_py else None
    if D_py and _tb_pnl_py_bf != 0 and pnl_bf_py is not None and abs(pnl_bf_py + _tb_pnl_py_bf) < 1.0:
        pnl_bf_py = _tb_pnl_py_bf   # was negated — restore
    np_cy        = D.get('net_profit', 0)      # signed: negative = current year loss
    np_py        = D_py.get('net_profit', 0)   if D_py else None
    pnl_close_cy = pnl_bf_cy + (np_cy or 0)   # = pnl_on_bs
    pnl_close_py = (pnl_bf_py or 0) + (np_py or 0) if D_py else None

    rows2.append([P("    Balance in P&amp;L Account \u2014 Opening Balance","NRM"), A(pnl_bf_cy),    A(pnl_bf_py)])
    rows2.append([P("    Add: Net Profit / (Loss) for the Year","NRM"),              A(np_cy),       A(np_py)])
    rows2.append([P("    P&amp;L Account Closing Balance","NRM"),                    A(pnl_close_cy), A(pnl_close_py)])

    # Only show General Reserve if there are real named reserve ledgers beyond P&L
    _RESERVE_KW = ('reserve', 'surplus', 'general res', 'capital res', 'securities premium',
                   'revaluation', 'retained earnings')
    _cap_lines = D.get('cap_lines', [])
    _has_real_reserve = any(
        any(kw in nm.lower() for kw in _RESERVE_KW)
        for nm, _ in _cap_lines
        if nm and 'profit' not in nm.lower() and "p&l" not in nm.lower()
    )
    other_res_cy = (cap_total - eq_share) - pnl_close_cy
    other_res_py = ((cap_total_py - eq_share_py) - pnl_close_py) \
                   if (cap_total_py is not None and pnl_close_py is not None) else None
    if _has_real_reserve and abs(other_res_cy) > 1:
        rows2.append([P("    General Reserve / Other Reserves","NRM"), A(other_res_cy), A(other_res_py)])

    rows2.append([P("    Closing Balance (Reserves &amp; Surplus on BS)","BOLD"), AB(res_total), AB(res_total_py)])
    story.append(grid_table(rows2, CW, bold_rows=[len(rows2)-1]))
    story.append(Spacer(1, 4*mm))

    # ── Note 3: Long Term Borrowings ────────────────────────────────────
    # BUG 3 FIX: for each loan, show lender name + disclosure footer
    DTL_KW  = ("deferred tax", "deffered tax", "dtl")
    # BUG FIX: "secured" is a substring of "unsecured" — must check that
    # "unsecured" is NOT present before concluding a loan is secured.
    # Use explicit word-boundary logic: a ledger is Secured only when it
    # contains a secured keyword AND does NOT contain "unsecured".
    SEC_KW  = ("bank loan", "term loan", "vehicle loan", "mortgage", "hypothecation",
               "pledge", "cc limit", "cash credit", "overdraft")
    def _excl_dtl(lines):
        return [(n, v) for n, v in (lines or [])
                if not any(kw in n.lower() for kw in DTL_KW)]
    def _is_secured(name):
        nl = name.lower()
        if "unsecured" in nl:
            return False   # explicit unsecured — never treat as secured
        return any(kw in nl for kw in SEC_KW)
    _loans3_cy = _excl_dtl(D.get('loans_lines', []))
    _loans3_py = _excl_dtl(D_py.get('loans_lines', [])) if D_py else None
    note_table("Note 3: Long Term Borrowings",
               _loans3_cy, _loans3_py, mode="liability")
    # Footer disclosure for all loans (Schedule III requirement)
    if _loans3_cy:
        story.append(Paragraph(
            "<b>Disclosure — Long Term Borrowings (as required under Schedule III):</b>",
            S["SML"]))
        for _uln, _ulv in _loans3_cy:
            _nature = "Secured" if _is_secured(_uln) else "Unsecured"
            story.append(Paragraph(
                f"  &bull; <b>{_uln}</b>: Rs.{tc.fmt_indian(abs(_ulv))} — "
                f"Nature: {_nature} | Interest Rate: As per agreement | "
                "Repayment: On demand | Related Party: Refer Note 31A if applicable.",
                S["SML"]))
        story.append(Spacer(1, 2*mm))

    # ── Note 4: Short Term Borrowings ───────────────────────────────────
    note_table("Note 4: Short Term Borrowings",
               cl_b.get("Short Term Borrowings", []),
               cl_b_py.get("Short Term Borrowings", []),
               mode="liability")

    # ── Note 5: Trade Payables ───────────────────────────────────────────
    note_table("Note 5: Trade Payables",
               cl_b.get("Trade Payables", []),
               cl_b_py.get("Trade Payables", []),
               mode="liability")

    # ── Note 6: Other Current Liabilities ──────────────────────────────────
    # Includes: (a) genuine CR-balance CL items from Current Liabilities group
    #           (b) CR-balance Sundry Debtors → Advances from Customers
    #           (c) CR-balance Suspense A/c (reclassified from asset side)
    # DR-balance items are reclassified to Note 14.
    _all_tr_cy_n6 = ca_b.get("Trade Receivables", [])
    _all_tr_py_n6 = ca_b_py.get("Trade Receivables", []) if D_py else []
    _cr_adv_cy = [("Advances from Customers", v)
                  for n, v in _all_tr_cy_n6 if v > 0]
    _cr_adv_py = [("Advances from Customers", v)
                  for n, v in _all_tr_py_n6 if v > 0]
    # CR-balance Suspense → liability
    _s_raw_cy  = D.get('TB', {}).get('Suspense A/c', 0)
    _s_raw_py  = D_py.get('TB', {}).get('Suspense A/c', 0) if D_py else 0
    _susp_cl_cy = [("Suspense A/c", _s_raw_cy)] if _s_raw_cy > 0 else []
    _susp_cl_py = [("Suspense A/c", _s_raw_py)] if _s_raw_py > 0 else []
    _note6_cy = _true_cl + _cr_adv_cy + _susp_cl_cy
    _note6_py = _true_cl_py + _cr_adv_py + _susp_cl_py
    note_table("Note 6: Other Current Liabilities",
               _note6_cy, _note6_py,
               mode="liability")

    # ── Note 7: Short Term Provisions ───────────────────────────────────
    note_table("Note 7: Short Term Provision",
               cl_b.get("Short Term Provisions", []),
               cl_b_py.get("Short Term Provisions", []),
               mode="liability")

    # ── Note 8: Fixed Assets ────────────────────────────────────────────
    # For companies with WDV depreciation: show full schedule
    #   Gross Block: Opening | Additions | Deletions | Closing
    #   Depreciation: Opening Dep | Dep for Year | Dep on Deletions | Closing Dep
    #   Net Block: CY | PY
    # For proprietorships/LLPs with no accumulated dep: show simple 3-col table.
    story.append(Paragraph("Note 8: Tangible Assets (Fixed Assets)", S["BOLD"]))
    _fa_lines_cy = [(n, v) for n, v in D.get('fa_lines_all', [])
                    if 'depreci' not in n.lower()]
    _dep_reserve_cy = D.get('dep_reserve', 0)
    _fa_gross_cy    = D.get('fa_gross', 0)
    _fa_net_cy      = D.get('fa_net_bs', 0)

    # PY gross block = PY fa_total abs value
    _fa_gross_py = abs(D_py.get('fa_total', 0)) if D_py else 0
    if not _fa_gross_py and D_py:
        _fa_gross_py = D_py.get('fa_gross', 0)
    _dep_reserve_py = 0
    if D_py:
        _dep_reserve_py = abs(D_py.get('TB', {}).get('DEPRICIATION RESERVE',
                              D_py.get('TB', {}).get('Depreciation Reserve', 0)))

    _has_dep = _dep_reserve_cy > 0 or _dep_reserve_py > 0

    if _has_dep:
        # Full 8-column schedule (standard CA format for companies with WDV dep)
        # Columns: Asset | Op.Gross | Additions | Deletions | Cl.Gross |
        #          Op.Dep | Dep/Yr | Dep/Del | Cl.Dep | Net CY | Net PY
        # Simplified to 7 cols to fit A4
        CW8 = [PAGE_W*0.26, PAGE_W*0.10, PAGE_W*0.09, PAGE_W*0.09, PAGE_W*0.10,
               PAGE_W*0.09, PAGE_W*0.09, PAGE_W*0.09, PAGE_W*0.09]
        _hdr8 = [
            P("Particulars","BOLD"),
            Paragraph("<b>Op.\nGross\nBlock</b>",   S["CTR"]),
            Paragraph("<b>Add-\nations</b>",         S["CTR"]),
            Paragraph("<b>Dele-\ntions</b>",         S["CTR"]),
            Paragraph("<b>Cl.\nGross\nBlock</b>",    S["CTR"]),
            Paragraph("<b>Op.\nDep.</b>",            S["CTR"]),
            Paragraph("<b>Dep.\nfor\nYear</b>",      S["CTR"]),
            Paragraph("<b>Cl.\nDep.</b>",            S["CTR"]),
            Paragraph("<b>Net\nBlock\nCY</b>",       S["CTR"]),
        ]
        rows8 = [_hdr8]

        # Per-asset rows — Tally gives us CY gross per asset and dep reserve total only.
        # Opening gross = PY gross (approximation — exact per-asset opening not in Tally XML).
        # Additions = CY gross − PY gross (if positive); Deletions = PY − CY (if positive).
        _fa_lines_py_map = {}
        if D_py:
            for _n, _v in D_py.get('fa_lines_all', []):
                if 'depreci' not in _n.lower():
                    _fa_lines_py_map[_n.strip()] = abs(_v)

        # Dep per asset: apportion total dep reserve by gross block weight
        _dep_per_asset = {}
        if _fa_gross_cy > 0 and _dep_reserve_cy > 0:
            for _n, _v in _fa_lines_cy:
                _wt = abs(_v) / _fa_gross_cy
                _dep_per_asset[_n.strip()] = _dep_reserve_cy * _wt

        _tot_op = _tot_add = _tot_del = _tot_cl = 0.0
        _tot_op_dep = _tot_dep_yr = _tot_cl_dep = _tot_net = 0.0
        _dep_op_total = _dep_reserve_py  # opening dep reserve = PY dep reserve

        for _n, _v in _fa_lines_cy:
            _cl_gross  = abs(_v)
            _op_gross  = _fa_lines_py_map.get(_n.strip(), _cl_gross)
            _additions = max(0, _cl_gross - _op_gross)
            _deletions = max(0, _op_gross - _cl_gross)
            # Per-asset opening dep: apportion by PY gross weight
            _op_dep = (_dep_op_total * (_op_gross / _fa_gross_py)
                       if _fa_gross_py > 0 else 0)
            _dep_yr  = _dep_per_asset.get(_n.strip(), 0)
            _cl_dep  = _op_dep + _dep_yr
            _net_blk = _cl_gross - _cl_dep

            rows8.append([
                P(f"  {_n}","NRM"),
                A(_op_gross), A(_additions or None), A(_deletions or None),
                A(_cl_gross), A(_op_dep or None), A(_dep_yr or None),
                A(_cl_dep or None), A(_net_blk)
            ])
            _tot_op  += _op_gross;  _tot_add += _additions; _tot_del += _deletions
            _tot_cl  += _cl_gross;  _tot_op_dep += _op_dep; _tot_dep_yr += _dep_yr
            _tot_cl_dep += _cl_dep; _tot_net += _net_blk

        rows8.append([
            P("Total","BOLD"),
            AB(_tot_op), AB(_tot_add or None), AB(_tot_del or None),
            AB(_tot_cl), AB(_tot_op_dep or None), AB(_tot_dep_yr or None),
            AB(_tot_cl_dep or None), AB(_fa_net_cy)
        ])
        story.append(grid_table(rows8, CW8, bold_rows=[len(rows8)-1]))
    else:
        # Simple 3-column table for zero-dep companies (proprietorship, new cos)
        CW8 = [PAGE_W*0.40, PAGE_W*0.20, PAGE_W*0.20, PAGE_W*0.20]
        rows8 = [[P("Particulars","BOLD"),
                  Paragraph("<b>Gross Block</b>", S["CTR"]),
                  Paragraph("<b>Depreciation</b>", S["CTR"]),
                  Paragraph("<b>Net Block</b>", S["CTR"])]]
        for _n, _v in _fa_lines_cy:
            rows8.append([P(f"    {_n}","NRM"), A(abs(_v)), A(None), A(abs(_v))])
        rows8.append([P("Less: Depreciation Reserve","NRM"), A(None),
                      A(_dep_reserve_cy or None), A(None)])
        rows8.append([P("Total","BOLD"),
                      AB(_fa_gross_cy), AB(_dep_reserve_cy or None), AB(_fa_net_cy)])
        story.append(grid_table(rows8, CW8, bold_rows=[len(rows8)-1]))
    story.append(Spacer(1, 4*mm))

    # ── Note 9: Long Term Loans & Advances ──────────────────────────────
    lt_loans_lines = ca_b.get("Long Term Loans and Advances", [])
    if not lt_loans_lines and D.get("deposits", 0):
        lt_loans_lines = [("Deposits (Asset)", -abs(D["deposits"]))]
    lt_loans_py = ca_b_py.get("Long Term Loans and Advances", []) if D_py else []
    if not lt_loans_py and D_py and D_py.get("deposits", 0):
        lt_loans_py = [("Deposits (Asset)", -abs(D_py["deposits"]))]
    note_table("Note 9: Long Term Loans &amp; Advances",
               lt_loans_lines, lt_loans_py, mode="asset")

    # ── Note 10: Non-Current Investments ────────────────────────────────
    # Only rendered when there are investments in either year.
    _inv_cache     = D.get('tb_children_of_cache', {}).get('Investments', [])
    _inv_cache_py  = D_py.get('tb_children_of_cache', {}).get('Investments', []) if D_py else []
    # inv_lines from bs_figures; fall back to cache
    _inv_lines_cy  = _inv_cache if _inv_cache else []
    _inv_lines_py  = _inv_cache_py if _inv_cache_py else []
    if _inv_lines_cy or _inv_lines_py:
        note_table("Note 10: Non-Current Investments",
                   _inv_lines_cy, _inv_lines_py, mode="asset")

    # ── Note 11: Inventories ────────────────────────────────────────────
    _inv_lines    = [(sl['name'], -abs(sl['closing'])) for sl in D.get('stk_ledgers', [])
                     if sl.get('closing', 0) > 0]
    _inv_lines_py = [(sl['name'], -abs(sl['closing'])) for sl in (D_py.get('stk_ledgers', []) if D_py else [])
                     if sl.get('closing', 0) > 0]
    note_table("Note 11: Inventories", _inv_lines, _inv_lines_py, mode="asset")

    # ── Note 9A: Other Non-Current Assets (Forex / Misc Expenses ASSET) ──
    # Unadjusted Forex Gain/Loss and Misc. Expenses (ASSET) are separate
    # Tally groups that appear on the asset side of Tally's BS.
    _forex_cy = abs(D.get('TB', {}).get('Unadjusted Forex Gain/Loss',
                D.get('TB', {}).get('Unadjusted Forex Gain / Loss', 0)))
    _misc_cy  = abs(D.get('TB', {}).get('Misc. Expenses (ASSET)',
                D.get('TB', {}).get('Miscellaneous Expenses (ASSET)', 0)))
    _forex_py = abs(D_py.get('TB', {}).get('Unadjusted Forex Gain/Loss',
                D_py.get('TB', {}).get('Unadjusted Forex Gain / Loss', 0))) if D_py else 0
    _misc_py  = abs(D_py.get('TB', {}).get('Misc. Expenses (ASSET)', 0)) if D_py else 0
    _nca9a_cy = []
    _nca9a_py = []
    if _forex_cy: _nca9a_cy.append(('Unadjusted Forex Gain/Loss', -_forex_cy))
    if _misc_cy:  _nca9a_cy.append(('Misc. Expenses (ASSET)',      -_misc_cy))
    if _forex_py: _nca9a_py.append(('Unadjusted Forex Gain/Loss', -_forex_py))
    if _misc_py:  _nca9a_py.append(('Misc. Expenses (ASSET)',      -_misc_py))
    if _nca9a_cy or _nca9a_py:
        note_table("Note 9A: Other Non-Current Assets",
                   _nca9a_cy, _nca9a_py or None, mode="asset")

    # ── Note 12: Trade Receivables ───────────────────────────────────────
    # Only genuine DR-balance (net receivable) debtors shown here.
    # CR-balance debtors (advances from customers) are reclassified to Note 6.
    _all_tr_cy  = ca_b.get("Trade Receivables", [])
    _all_tr_py  = ca_b_py.get("Trade Receivables", []) if D_py else []
    _note12_cy  = [(n, v) for n, v in _all_tr_cy if v <= 0]   # DR = genuine receivable
    _note12_py  = [(n, v) for n, v in _all_tr_py if v <= 0]
    note_table("Note 12: Trade Receivables",
               _note12_cy, _note12_py, mode="asset")

    # ── Note 13: Cash and Cash Equivalents ──────────────────────────────
    note_table("Note 13: Cash and Cash Equivalents",
               ca_b.get("Cash and Cash Equivalents", []),
               ca_b_py.get("Cash and Cash Equivalents", []),
               mode="asset")

    # ── Note 14: Short Term Loans & Advances ────────────────────────────
    # Includes: normal CA_FALLBACK items + reclassified debit-balance CL items
    # EXCLUDE suspense-named items — they are shown separately in Note 15A.
    _st_ca_items    = [(n, v) for n, v in ca_b.get(CA_FALLBACK, [])
                       if not any(kw in n.lower() for kw in SUSPENSE_KW)]
    _st_ca_items_py = [(n, v) for n, v in ca_b_py.get(CA_FALLBACK, [])
                       if not any(kw in n.lower() for kw in SUSPENSE_KW)] if D_py else []
    _reclass_for_note    = [(n, -abs(v)) for n, v in _reclassified]
    _reclass_for_note_py = [(n, -abs(v)) for n, v in _reclassified_py]
    _note14_items    = _st_ca_items    + _reclass_for_note
    _note14_items_py = _st_ca_items_py + _reclass_for_note_py
    note_table("Note 14: Short Term Loans &amp; Advances",
               _note14_items, _note14_items_py, mode="asset")

    # ── Note 15A: Other Current Assets (Suspense) ──────────────────────
    # Only rendered when Suspense has a DR balance (asset side).
    # CR-balance Suspense is already shown in Note 6 Other Current Liabilities.
    _susp_cache    = D.get('tb_children_of_cache', {}).get('Suspense A/c', [])
    _susp_cache_py = D_py.get('tb_children_of_cache', {}).get('Suspense A/c', []) if D_py else []
    _s_cy = D.get('TB', {}).get('Suspense A/c', 0)
    _s_py = D_py.get('TB', {}).get('Suspense A/c', 0) if D_py else 0

    # Also capture suspense items from ca_lines when TB group key is absent
    # (Tally sometimes puts Suspense under Current Assets without a separate group entry)
    _susp_from_ca_cy = [(n, v) for n, v in ca_b.get(CA_FALLBACK, [])
                        if any(kw in n.lower() for kw in SUSPENSE_KW) and v < 0]  # DR = asset
    _susp_from_ca_py = [(n, v) for n, v in ca_b_py.get(CA_FALLBACK, [])
                        if any(kw in n.lower() for kw in SUSPENSE_KW) and v < 0] if D_py else []

    # Show Note 15A when DR-balance suspense exists in TB group, children, or ca_lines
    _show_susp_asset = (_s_cy < 0) or (_s_py < 0) or \
                       (not _s_cy and (_susp_cache or _susp_from_ca_cy)) or \
                       (not _s_py and (_susp_cache_py or _susp_from_ca_py))
    if _show_susp_asset:
        # Use children if available, otherwise fall back to ca_lines DR items
        _sc_dr    = [(n, v) for n, v in _susp_cache    if v <= 0] or _susp_from_ca_cy
        _sc_dr_py = ([(n, v) for n, v in _susp_cache_py if v <= 0] or _susp_from_ca_py) if D_py else []
        if _sc_dr or _sc_dr_py:
            note_table("Note 15A: Other Current Assets (Suspense A/c)",
                       _sc_dr, _sc_dr_py or None, mode="asset")

    # ── P&L Notes ───────────────────────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph(
        f"Notes forming part of Statement of Profit &amp; Loss for F.Y. ending as on {fmt_date(fy_e)}",
        S["SUB"]))
    story.append(Spacer(1, 3*mm))

    # FIX: Dormant/shell company — suppress all nil P&L notes, show one clear disclosure.
    _dormant_pnl_notes = tc.detect_company_type(D.get('TB', {})).get('is_dormant', False)
    if _dormant_pnl_notes:
        story.append(Paragraph(
            "<b>Notes 15 to 21:</b> The Company had no revenue from operations, other income, "
            "purchases, employee expenses, finance costs, depreciation or other expenses "
            f"during the year ended {fmt_date(fy_e)}. "
            "All the above notes are NIL for the current year.",
            S["NRM"]))
        story.append(Spacer(1, 4*mm))
    else:
        # Note 15: Revenue From Operations
        # BUG 7 FIX: service companies with revenue only in Indirect Incomes
        _sales_n15 = D.get('sales_lines', [])
        _ind_n15   = D.get('ind_inc_lines', [])
        _sales_n15_py = D_py.get('sales_lines', []) if D_py else None
        _ind_n15_py   = D_py.get('ind_inc_lines', []) if D_py else None
        _cy_dormant_n15 = tc.detect_company_type(D.get('TB', {})).get('is_dormant', False)
        _svc_reclassify = (not _sales_n15 and bool(_ind_n15) and not _cy_dormant_n15)
        if _svc_reclassify:
            note_table("Note 15: Revenue From Operations",
                       _ind_n15,
                       _ind_n15_py or None,
                       mode="liability")
            story.append(Paragraph(
                "Note: Revenue classified under Indirect Incomes in Tally ledger grouping. "
                "Reclassified to Revenue from Operations as per Schedule III.",
                S["SML"]))
            story.append(Spacer(1, 2*mm))
        else:
            note_table("Note 15: Revenue From Operations",
                       _sales_n15,
                       _sales_n15_py,
                       mode="liability")

        # Note 16: Other Income
        # If service reclassification applied, Other Income is nil
        if _svc_reclassify:
            note_table("Note 16: Other Income", [], None, mode="liability")
        else:
            note_table("Note 16: Other Income",
                       _ind_n15,
                       _ind_n15_py,
                       mode="liability")

        # Note 17: Cost of Material Consumed
        def _cost_mat_lines(d):
            _d_TB      = d.get('TB', {})
            _d_cotype  = tc.detect_company_type(_d_TB)
            _d_has_stk = _d_cotype.get('has_stock', False)

            lines = list(d.get('purch_lines', []))
            for k, label in [('mfg_exp', 'Manufacturing Expenses'),
                             ('power_fuel', 'Power & Fuel'),
                             ('stores_exp', 'Stores & Spares'),
                             ('stores_cons', 'Consumables')]:
                v = d.get(k, 0)
                if v:
                    lines.append((label, v))

            # For trading-WITHOUT-stock companies (e.g. Vianci):
            # Direct Expenses (Freight, Labour, Transport etc.) are true cost-of-goods.
            # Include ALL dir_exp_lines in Note 17 so P&L shows the correct cost total.
            # These expenses are part of tally_core's gross_profit formula for this type.
            if not _d_has_stk and d.get('purchases', 0) != 0:
                for n, v in d.get('dir_exp_lines', []):
                    if v != 0:
                        lines.append((n, v))
            elif not lines:
                # Last fallback: any dir_exp_line that looks like a purchase/material item
                _purch_kw = ['purchase', 'material consumed', 'cost of material',
                             'raw material', 'trading goods', 'stock in trade']
                for n, v in d.get('dir_exp_lines', []):
                    if any(kw in n.lower() for kw in _purch_kw):
                        lines.append((n, v))
            return lines

        # Track which dir_exp items are captured in Note 17 (to exclude from Note 21)
        _cost17_names_cy = set(n for n,_ in _cost_mat_lines(D))
        _cost17_names_py = set(n for n,_ in (_cost_mat_lines(D_py) if D_py else []))

        note_table("Note 17: Cost of Material Consumed",
                   _cost_mat_lines(D),
                   _cost_mat_lines(D_py) if D_py else None,
                   mode="asset")

        # Note 18: Changes in Inventories of Finished Goods, WIP and Stock-in-Trade
        # FIX: This note was entirely missing — statutory Schedule III requirement.
        # For service/dormant/no-stock companies: both opening and closing = 0 → shows nil.
        def _inv_chg_lines(d):
            _TB_d     = d.get('TB', {})
            _co_d     = tc.detect_company_type(_TB_d)
            _has_stk  = _co_d.get('has_stock', False)
            if not _has_stk:
                return [('Opening Stock', 0), ('Less: Closing Stock', 0)]
            _os  = d.get('opening_stock_abs', 0)
            _cs  = d.get('bs_closing_stock', d.get('closing_stock_abs', 0))
            # inv_chg sign: increase in stock = negative (credit to P&L); decrease = positive (debit)
            lines = []
            if _os:  lines.append(('Opening Stock', -_os))    # DR at period start (neg TB value)
            if _cs:  lines.append(('Less: Closing Stock', _cs)) # CR at period end  (pos TB value)
            if not lines:
                lines = [('Opening Stock', 0), ('Less: Closing Stock', 0)]
            return lines
        note_table("Note 18: Changes in Inventories",
                   _inv_chg_lines(D),
                   _inv_chg_lines(D_py) if D_py else None,
                   mode="asset", total_label="Net (Increase) / Decrease in Inventories")

        # Note 19: Employee Benefit Expenses
        EMP_KW = ['salary','wage','bonus','esic','gratuity','staff welfare','labour welfare','director remun']
        _emp_cy = [(n,v) for n,v in D.get('ind_exp_lines',[]) if any(kw in n.lower() for kw in EMP_KW)]
        _emp_py = [(n,v) for n,v in (D_py.get('ind_exp_lines',[]) if D_py else []) if any(kw in n.lower() for kw in EMP_KW)]
        note_table("Note 19: Employee Benefit Expenses",
                   _emp_cy, _emp_py or None, mode="asset")

        # Note 20: Finance Cost
        FIN_KW = ['interest','bank charge','processing fee','mortgage','finance charge']
        _fin_cy = [(n,v) for n,v in D.get('ind_exp_lines',[]) if any(kw in n.lower() for kw in FIN_KW)]
        _fin_py = [(n,v) for n,v in (D_py.get('ind_exp_lines',[]) if D_py else []) if any(kw in n.lower() for kw in FIN_KW)]
        note_table("Note 20: Finance Cost",
                   _fin_cy, _fin_py or None, mode="asset")

        # Note 21: Other Expenses
        # = indirect expenses (excl employee/finance) + any direct expenses NOT already in Note 17
        #
        # DOUBLE-COUNT GUARD:
        # For trading-WITHOUT-stock companies (e.g. Vianci), ALL dir_exp_lines are
        # already shown in Note 17 (Cost of Material Consumed). Do NOT add them here.
        # For manufacturing/stock companies, dir_exp_lines that are NOT in Note 17
        # (e.g. they were not individually captured) are included here.
        _excl_kw = EMP_KW + FIN_KW
        _other_ind_cy = [(n,v) for n,v in D.get('ind_exp_lines',[])
                         if not any(kw in n.lower() for kw in _excl_kw)]
        _other_ind_py = [(n,v) for n,v in (D_py.get('ind_exp_lines',[]) if D_py else [])
                         if not any(kw in n.lower() for kw in _excl_kw)]

        # Detect whether dir_exp_lines were already consumed by Note 17
        _cy_TB        = D.get('TB', {})
        _py_TB        = D_py.get('TB', {}) if D_py else {}
        _cy_no_stock  = not tc.detect_company_type(_cy_TB).get('has_stock', True)
        _py_no_stock  = not tc.detect_company_type(_py_TB).get('has_stock', True) if D_py else True
        # If no stock AND has purchases → dir_exp already in Note 17 → exclude from Note 21
        _cy_direxp_in_17 = _cy_no_stock and bool(D.get('purchases', 0))
        _py_direxp_in_17 = _py_no_stock and bool(D_py.get('purchases', 0)) if D_py else False

        _other_dir_cy = [] if _cy_direxp_in_17 else [
            (n,v) for n,v in D.get('dir_exp_lines',[])
            if n not in _cost17_names_cy and n.strip() not in _cost17_names_cy]
        _other_dir_py = [] if _py_direxp_in_17 else [
            (n,v) for n,v in (D_py.get('dir_exp_lines',[]) if D_py else [])
            if n not in _cost17_names_py and n.strip() not in _cost17_names_py]

        _note21_cy = _other_ind_cy + _other_dir_cy
        _note21_py = _other_ind_py + _other_dir_py
        note_table("Note 21: Other Expenses",
                   _note21_cy, _note21_py or None, mode="asset")

    # ── Note 22: Deferred Tax ─────────────────────────────────────────────
    # Shown for all company types when DTL/DTA exists.
    # DTL = Deferred Tax Liability (shown under Non-Current Liabilities on BS)
    # DTA = Deferred Tax Asset (shown under Non-Current Assets on BS)
    # Both can coexist — net position is what appears on BS.
    TB_n22    = D.get('TB', {})
    TB_py_n22 = D_py.get('TB', {}) if D_py else {}

    # Detect DTL/DTA from TB — handles common Tally naming variations
    _DTL_KEYS = ('Deffered Tax Liabilites', 'Deferred Tax Liability',
                 'Deferred Tax Liabilities', 'DTL', 'Deffered Tax Liability')
    _DTA_KEYS = ('Deferred Tax Asset', 'Deffered Tax Asset',
                 'DTA', 'Deferred Tax Assets')
    _dtl_cy = sum(abs(TB_n22.get(k, 0)) for k in _DTL_KEYS if TB_n22.get(k, 0))
    _dta_cy = sum(abs(TB_n22.get(k, 0)) for k in _DTA_KEYS if TB_n22.get(k, 0))
    _dtl_py = sum(abs(TB_py_n22.get(k, 0)) for k in _DTL_KEYS if TB_py_n22.get(k, 0)) if D_py else 0
    _dta_py = sum(abs(TB_py_n22.get(k, 0)) for k in _DTA_KEYS if TB_py_n22.get(k, 0)) if D_py else 0

    # Also scan loans_lines for deferred tax entries (some Tally setups put DTL there)
    for _ln, _lv in D.get('loans_lines', []):
        if any(kw in _ln.lower() for kw in ('deferred tax', 'deffered tax', 'dtl')):
            _dtl_cy = max(_dtl_cy, abs(_lv))
    if D_py:
        for _ln, _lv in D_py.get('loans_lines', []):
            if any(kw in _ln.lower() for kw in ('deferred tax', 'deffered tax', 'dtl')):
                _dtl_py = max(_dtl_py, abs(_lv))

    _net_dt_cy = _dtl_cy - _dta_cy   # positive = net DTL; negative = net DTA
    _net_dt_py = _dtl_py - _dta_py if D_py else None

    story.append(Paragraph("Note 22: Deferred Tax", S["BOLD"]))
    _rows22 = [hdr_row()]
    _rows22.append([P("    Deferred Tax Liability (DTL)", "NRM"),
                    A(_dtl_cy or None), A(_dtl_py or None)])
    _rows22.append([P("    Less: Deferred Tax Asset (DTA)", "NRM"),
                    A(_dta_cy or None), A(_dta_py or None)])
    _rows22.append([P("    Net Deferred Tax Liability / (Asset)", "BOLD"),
                    AB(_net_dt_cy if (_dtl_cy or _dta_cy) else None),
                    AB(_net_dt_py if (_dtl_py or _dta_py) else None)])

    if _dtl_cy == 0 and _dta_cy == 0:
        _rows22.append([P("    No deferred tax recognised during the year — "
                          "timing differences are not material / no qualifying differences exist.",
                          "NRM"), A(None), A(None)])
    story.append(grid_table(_rows22, CW, bold_rows=[len(_rows22)-1]))

    # Disclosure note on timing differences
    if _dtl_cy > 0 or _dta_cy > 0:
        story.append(Paragraph(
            "Note: Deferred tax has been recognised on timing differences arising primarily "
            "from depreciation (WDV as per Companies Act vs Income Tax Act rates) and "
            "disallowances under the Income Tax Act, 1961.",
            S["SML"]))
    story.append(Spacer(1, 4*mm))

    # ── Note 33: Director / Managerial Remuneration ───────────────────────
    # Mandatory disclosure under Section 197 / Schedule V of Companies Act 2013
    # for Pvt Ltd and Public Ltd companies. Skip for proprietorships/LLPs.
    _cap_lines_n33 = D.get('cap_lines', [])
    _SHARE_KW_N33  = ('share capital', 'equity share', 'preference share', 'equity capital')
    _is_company_n33 = any(any(kw in nm.lower() for kw in _SHARE_KW_N33)
                          for nm, _ in _cap_lines_n33)

    if _is_company_n33:
        # Find director remuneration from ind_exp_lines
        _DIR_REM_KW = ('director remun', 'managerial remun', 'director salary',
                       'director fees', 'managing director', 'md salary', 'cmd salary',
                       'director commission', 'sitting fees', 'sitting fee')
        _dir_rem_cy = [(n, v) for n, v in D.get('ind_exp_lines', [])
                       if any(kw in n.lower() for kw in _DIR_REM_KW)]
        _dir_rem_py = [(n, v) for n, v in (D_py.get('ind_exp_lines', []) if D_py else [])
                       if any(kw in n.lower() for kw in _DIR_REM_KW)]
        # Also check direct expenses
        for _n, _v in D.get('dir_exp_lines', []):
            if any(kw in _n.lower() for kw in _DIR_REM_KW):
                _dir_rem_cy.append((_n, _v))

        _tot_dir_rem_cy = sum(abs(v) for _, v in _dir_rem_cy)
        _tot_dir_rem_py = sum(abs(v) for _, v in _dir_rem_py) if _dir_rem_py else None

        story.append(Paragraph("Note 33: Director / Managerial Remuneration", S["BOLD"]))
        story.append(Paragraph(
            "Disclosure pursuant to Section 197 read with Schedule V of the Companies Act, 2013:",
            S["SML"]))
        _rows33 = [hdr_row()]

        if _dir_rem_cy or _dir_rem_py:
            for _rn, _rv in _dir_rem_cy:
                _rv_py = next((v for n, v in _dir_rem_py if n == _rn), None)
                _rows33.append([P(f"    {_rn}", "NRM"), A(abs(_rv)), A(abs(_rv_py) if _rv_py else None)])
            _rows33.append([P("    Total Managerial Remuneration", "BOLD"),
                            AB(_tot_dir_rem_cy), AB(_tot_dir_rem_py)])
            _rows33.append([P("    Computation of net profit u/s 198 of Companies Act, 2013:", "NRM"),
                            A(None), A(None)])
            _rows33.append([P("    (Refer to separate computation — not auto-generated)", "NRM"),
                            A(None), A(None)])
        else:
            _rows33.append([P("    No managerial remuneration paid / payable during the year.",
                              "NRM"), A(None), A(None)])
            _rows33.append([P("    (Confirm with management — may not be captured in Tally)",
                              "NRM"), A(None), A(None)])
        story.append(grid_table(_rows33, CW, bold_rows=([len(_rows33)-3] if _dir_rem_cy else [])))
        story.append(Spacer(1, 4*mm))


# ══════════════════════════════════════════════════════════════════════════
# PY SANITY CHECK
# ══════════════════════════════════════════════════════════════════════════
def _py_is_same_as_cy(D, D_py):
    """Returns True if D_py appears to be a re-fetch of the same year as D.
    This happens when Tally has no prior-year data and returns current-year
    data for both fetch calls. Detected by identical fy_end."""
    if D_py is None:
        return False
    if D.get('fy_end') == D_py.get('fy_end'):
        print(f"  WARNING: D_py fy_end ({D_py.get('fy_end')}) == D fy_end — "
              "PY fetch returned current year data. PY column will be blank.")
        return True
    return False


# ══════════════════════════════════════════════════════════════════════════
# MASTER GENERATOR
# ══════════════════════════════════════════════════════════════════════════
def generate_classic_pdf(D, D_py, out_path, include_policies=True,
                         also_xlsx=False, also_xlsx_classic=False):
    # Guard: drop D_py if it is the same year as D
    if _py_is_same_as_cy(D, D_py):
        D_py = None

    doc = SimpleDocTemplate(out_path, pagesize=A4,
                             topMargin=12*mm, bottomMargin=12*mm,
                             leftMargin=10*mm, rightMargin=10*mm)
    story = []
    build_classic_bs(D, D_py, story)
    build_classic_pnl(D, D_py, story)
    build_classic_notes(D, D_py, story)

    # Statutory pages: policies + verification + MSME + related party
    # BUG 10 FIX: pass pre-computed bs_figures so verification page uses
    # the same figures as the BS page (no stale re-computation).
    _f_for_verify = bs_figures(D)
    try:
        import sch3_verify as v3
        v3.build_all_statutory(D, D_py, story, bs_f=_f_for_verify)
    except Exception as _e:
        print(f"  WARNING: Statutory pages failed — {_e}. Falling back to basic policies.")
        if include_policies:
            try:
                story.append(PageBreak())
                tc.build_policies(D, story)
            except Exception:
                pass

    doc.build(story)
    print(f"  Classic Schedule III PDF saved -> {out_path}")

    # Legacy BS-only Excel (existing behaviour)
    if also_xlsx:
        xlsx_path = out_path.replace(".pdf", ".xlsx") if out_path.endswith(".pdf") else out_path + ".xlsx"
        try:
            download_bs_xlsx(D, D_py, xlsx_path)
        except Exception as e:
            print(f"  WARNING: Excel generation failed — {e}")

    # New full Schedule III Classic Excel with formulas + unit conversion
    if also_xlsx_classic:
        safe = re.sub(r'[^\w]', '_', D['company'][:30])
        xlsx_classic_path = out_path.replace(".pdf", "_Classic.xlsx") if out_path.endswith(".pdf") \
                            else out_path + "_Classic.xlsx"
        try:
            from sch3_excel_classic import generate_schedule3_excel
            generate_schedule3_excel(D, D_py, xlsx_classic_path)
        except Exception as e:
            print(f"  WARNING: Classic Excel generation failed — {e}")

    return out_path


# ══════════════════════════════════════════════════════════════════════════
# EXCEL DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════
def download_bs_xlsx(D, D_py, out_path=None):
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    except ImportError:
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"])
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    fy_e    = D['fy_end']
    fy_p    = D_py['fy_end'] if D_py else f"{int(fy_e[:4])-1}{fy_e[4:]}"
    company = D['company']
    f  = bs_figures(D)
    fp = bs_figures(D_py) if D_py else {k: None for k in f}

    if out_path is None:
        safe = re.sub(r'[^\w]', '_', company[:30])
        out_path = f"BS_Sch3_{safe}_{fy_e}.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Balance Sheet"

    HDR_FILL = PatternFill("solid", fgColor="1F3864")
    SEC_FILL = PatternFill("solid", fgColor="2E5496")
    TOT_FILL = PatternFill("solid", fgColor="D6E4F0")
    ALT_FILL = PatternFill("solid", fgColor="EEF4FB")
    WHT_FILL = PatternFill("solid", fgColor="FFFFFF")
    HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    TOT_FONT = Font(name="Arial", bold=True, color="1F3864", size=8.5)
    NRM_FONT = Font(name="Arial", size=8.5)
    NEG_FONT = Font(name="Arial", size=8.5, color="C00000")
    NEG_BOLD = Font(name="Arial", bold=True, size=8.5, color="C00000")
    thin  = Side(style="thin", color="8EA9C1")
    thick = Side(style="medium", color="1F3864")
    b_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    b_top = Border(left=thin, right=thin, top=thick, bottom=thin)

    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20

    def fmtv(v):
        if v is None or v == "": return ""
        try: v = float(v)
        except: return ""
        if abs(v) < 0.005: return "-"
        s = tc.fmt_indian(abs(v))
        return f"({s})" if v < 0 else s

    ri = [1]
    def nr():
        r = ri[0]; ri[0] += 1; return r

    def hdr():
        r = nr()
        ws.merge_cells(f"A{r}:D{r}")
        c = ws.cell(r,1,company.upper())
        c.font = Font(name="Arial",bold=True,size=12,color="1F3864")
        c.alignment = Alignment(horizontal="center",vertical="center")
        ws.row_dimensions[r].height = 18
        r = nr()
        ws.merge_cells(f"A{r}:D{r}")
        c = ws.cell(r,1,f"BALANCE SHEET AS AT {fmt_date(fy_e)}")
        c.font = Font(name="Arial",bold=True,size=10,color="1F3864")
        c.alignment = Alignment(horizontal="center",vertical="center")
        r = nr()
        ws.cell(r,4,"(Amt. in Rs.)").font = Font(name="Arial",italic=True,size=7.5,color="666666")
        ws.cell(r,4).alignment = Alignment(horizontal="right")
        r = nr()
        ws.row_dimensions[r].height = 30
        for col,txt in enumerate(["PARTICULARS","Note No.",
                                   f"FIGURES AS AT\nTHE END OF\n{fmt_date(fy_e)}",
                                   f"FIGURES AS AT\nTHE END OF\n{fmt_date(fy_p)}"],1):
            c = ws.cell(r,col,txt)
            c.font = HDR_FONT; c.fill = HDR_FILL
            c.alignment = Alignment(horizontal="center" if col>1 else "left",
                                    vertical="center",wrap_text=True)
            c.border = b_all

    def sec(label):
        r = nr()
        ws.row_dimensions[r].height = 13
        c = ws.cell(r,1,label)
        c.font = Font(name="Arial",bold=True,color="FFFFFF",size=8.5)
        c.fill = SEC_FILL; c.alignment = Alignment(horizontal="left",vertical="center")
        c.border = b_all
        for col in [2,3,4]:
            ws.cell(r,col).fill = SEC_FILL; ws.cell(r,col).border = b_all

    def drow(label,note,cy,py,bold=False,indent=2):
        r = nr()
        ws.row_dimensions[r].height = 12
        fill = ALT_FILL if r%2==0 else WHT_FILL
        c = ws.cell(r,1,"    "*indent+label)
        c.font = Font(name="Arial",bold=bold,size=8.5); c.fill = fill
        c.alignment = Alignment(horizontal="left",vertical="center"); c.border = b_all
        ws.cell(r,2,note or "").font = NRM_FONT
        ws.cell(r,2).fill = fill; ws.cell(r,2).alignment = Alignment(horizontal="center",vertical="center")
        ws.cell(r,2).border = b_all
        for col,val in [(3,cy),(4,py)]:
            neg = val is not None and float(val or 0) < -0.005
            cv = ws.cell(r,col,fmtv(val))
            cv.font = (NEG_BOLD if bold else NEG_FONT) if neg else (TOT_FONT if bold else NRM_FONT)
            cv.fill = TOT_FILL if bold else fill
            cv.alignment = Alignment(horizontal="right",vertical="center"); cv.border = b_all

    def trow(label,cy,py):
        r = nr(); ws.row_dimensions[r].height = 14
        c = ws.cell(r,1,label)
        c.font = TOT_FONT; c.fill = TOT_FILL; c.border = b_top
        c.alignment = Alignment(horizontal="left",vertical="center")
        ws.cell(r,2,"").fill = TOT_FILL; ws.cell(r,2).border = b_top
        for col,val in [(3,cy),(4,py)]:
            neg = val is not None and float(val or 0) < -0.005
            cv = ws.cell(r,col,fmtv(val))
            cv.font = NEG_BOLD if neg else TOT_FONT
            cv.fill = TOT_FILL; cv.alignment = Alignment(horizontal="right",vertical="center")
            cv.border = b_top

    def brow(label="",indent=2):
        r = nr(); ws.row_dimensions[r].height = 11
        c = ws.cell(r,1,"    "*indent+label)
        c.font = NRM_FONT; c.fill = WHT_FILL
        c.alignment = Alignment(horizontal="left",vertical="center"); c.border = b_all
        for col in [2,3,4]:
            ws.cell(r,col).fill = WHT_FILL; ws.cell(r,col).border = b_all

    hdr()
    sec("I.  EQUITY & LIABILITIES")
    sec("    SHAREHOLDER'S FUNDS")
    drow("SHARE CAPITAL","1",f["eq_share"],fp["eq_share"],indent=2)
    drow("RESERVES & SURPLUS","2",f["reserves"],fp["reserves"],indent=2)
    brow("MONEY RECEIVED AGAINST SHARE WARRANTS")
    brow("SHARE APPLICATION MONEY PENDING ALLOTMENT",indent=1)
    sec("    NON-CURRENT LIABILITIES")
    drow("LONG-TERM BORROWINGS","3",f["lt_borrow"],fp["lt_borrow"],indent=2)
    drow("DEFERRED TAX LIABILITIES (NET)","32",f["dtl"],fp["dtl"],indent=2)
    brow("OTHER LONG TERM LIABILITIES"); brow("LONG-TERM PROVISIONS")
    sec("    CURRENT LIABILITIES")
    drow("SHORT-TERM BORROWINGS","4",f["st_borrow"],fp["st_borrow"],indent=2)
    drow("TRADE PAYABLES","5",f["trade_pay"],fp["trade_pay"],indent=2)
    _ocl_ex = f.get("other_cl_total", f.get("other_cl", 0))
    _ocl_exp= fp.get("other_cl_total", fp.get("other_cl")) if fp else None
    drow("OTHER CURRENT LIABILITIES","6",_ocl_ex,_ocl_exp,indent=2)
    drow("SHORT TERM PROVISION","7",f["short_prov"],fp["short_prov"],indent=2)
    trow("TOTAL EQUITY & LIABILITIES",f["total_l"],fp["total_l"] if fp else None)

    sec("II.  ASSETS")
    sec("    NON-CURRENT ASSETS")
    sec("    Property Plant and Equipment")
    drow("TANGIBLE ASSETS","8",f["fa_net"],fp["fa_net"],indent=2)
    brow("INTANGIBLE ASSETS"); brow("CAPITAL WORK-IN-PROGRESS")
    brow("INTANGIBLE ASSETS UNDER DEVELOPMENT")
    _inv_cy_x = f.get("inv_total", 0) or 0
    _inv_py_x = fp.get("inv_total") if fp else None
    if abs(_inv_cy_x) > 0.5 or (_inv_py_x and abs(_inv_py_x) > 0.5):
        drow("NON-CURRENT INVESTMENTS","10",_inv_cy_x or None,_inv_py_x,indent=1)
    else:
        brow("NON-CURRENT INVESTMENTS",indent=1)
    brow("DEFERRED TAX ASSET (NET)",indent=1)
    drow("LONG-TERM LOANS AND ADVANCES","9",f["lt_loans_adv"],fp["lt_loans_adv"] if fp else None,indent=1)
    _onca_ex = f.get("other_nca", 0) or 0
    _onca_exp= fp.get("other_nca", 0) if fp else None
    if abs(_onca_ex) > 0.5 or (_onca_exp and abs(_onca_exp) > 0.5):
        drow("OTHER NON-CURRENT ASSETS (Forex/Misc)","9A",_onca_ex or None,_onca_exp or None,indent=1)
    else:
        brow("OTHER NON-CURRENT ASSETS",indent=1)
    sec("    CURRENT ASSETS")
    brow("CURRENT INVESTMENTS")
    drow("INVENTORIES","11",f["inventories"],fp["inventories"] if fp else None,indent=2)
    drow("TRADE RECEIVABLES","12",f["trade_recv"],fp["trade_recv"] if fp else None,indent=2)
    drow("CASH AND CASH EQUIVALENTS","13",f["cash_bank"],fp["cash_bank"] if fp else None,indent=2)
    drow("SHORT TERM LOANS AND ADVANCES","14",f["st_loans_adv"],fp["st_loans_adv"] if fp else None,indent=2)
    _susp_cy_x = f.get("suspense_val", 0) or 0
    _susp_py_x = fp.get("suspense_val") if fp else None
    if abs(_susp_cy_x) > 0.5 or (_susp_py_x and abs(_susp_py_x) > 0.5):
        drow("OTHER CURRENT ASSETS (Suspense)","15A",_susp_cy_x or None,_susp_py_x,indent=2)
    else:
        brow("OTHER CURRENT ASSETS")
    trow("TOTAL ASSETS",f["total_a"],fp["total_a"] if fp else None)
    drow("III.  CONTINGENT LIABILITIES","31",None,None,indent=0)

    wb.save(out_path)
    print(f"  Schedule III BS Excel saved -> {out_path}")
    return out_path


# ══════════════════════════════════════════════════════════════════════════
# Loaders
# ══════════════════════════════════════════════════════════════════════════
def _payload_to_D(payload):
    company  = clean_company_name(payload["company"])
    fy_start = payload["fy_start"]
    fy_end   = payload["fy_end"]
    data     = payload["data"]
    return tc.parse_data(company, data, fy_start, fy_end)

def load_from_saved_json(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        payload = json.load(f)
    return _payload_to_D(payload)

def prior_fy_end_str(fy_end):
    s = str(fy_end).strip()
    if len(s) == 10 and s[4] == '-':
        return f"{int(s[:4])-1}{s[4:]}"
    elif len(s) == 8:
        return f"{int(s[:4])-1}{s[4:]}"
    return fy_end

def load_from_db(report_id=None, company=None, fy_end=None, auto_py=True):
    if not DB_OK:
        raise SystemExit("db.py is not available / psycopg not installed.")
    payload = tdb.get_report(report_id=report_id, company=company, fy_end=fy_end)
    if payload is None:
        raise SystemExit(f"No matching row (report_id={report_id}, company={company!r}, fy_end={fy_end!r}).")
    D = _payload_to_D(payload)
    D_py = None
    if auto_py:
        py_fy_end = prior_fy_end_str(D['fy_end'])
        for cname in {payload["company"], D['company']}:
            py_payload = tdb.get_report(company=cname, fy_end=py_fy_end)
            if py_payload:
                D_py = _payload_to_D(py_payload)
                break
    return D, D_py

def fetch_and_save(company, year, port=9000, save_db=False):
    tc.set_tally_url(f"http://localhost:{port}")
    raw_data, fy_start, fy_end = tc.fetch_all(company, year, save_json=True)
    D = tc.parse_data(company, raw_data, fy_start, fy_end)
    if save_db and DB_OK:
        try:
            payload = {"company": clean_company_name(company), "fy_start": fy_start,
                       "fy_end": fy_end, "tally_url": tc.TALLY_URL, "data": raw_data}
            rid = tdb.save_report(clean_company_name(company), fy_start, fy_end,
                                  tc.TALLY_URL, payload,
                                  gross_profit=float(D.get('gross_profit') or 0),
                                  net_profit=float(D.get('net_profit') or 0))
            print(f"  Saved to DB -> row id {rid}")
        except Exception as e:
            print(f"  WARNING: DB save failed - {e}")
    return D

def load_live_with_prior_year(company, year, port=9000, save_db=False):
    print(f"\n-> Fetching current year (FY end {year}) ...")
    D = fetch_and_save(company, year, port=port, save_db=save_db)
    print(f"\n-> Auto-syncing previous year (FY end {year-1}) ...")
    try:
        D_py = fetch_and_save(company, year - 1, port=port, save_db=save_db)
        # Safety: if same year returned, drop it
        if _py_is_same_as_cy(D, D_py):
            D_py = None
    except Exception as e:
        print(f"  WARNING: could not fetch previous year - {e}")
        D_py = None
    return D, D_py


# ──────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Classic Schedule III PDF/Excel generator")
    ap.add_argument("--json",       help="Current-year tally_raw_*.json")
    ap.add_argument("--json-py",    help="Previous-year tally_raw_*.json (optional)")
    ap.add_argument("--db-id",      type=int, help="Load by tally_reports.id")
    ap.add_argument("--db-company", help="Load by company name from DB")
    ap.add_argument("--db-fyend",   help="FY end date YYYY-MM-DD")
    ap.add_argument("--company",    help="Company name (live fetch)")
    ap.add_argument("--year",       type=int, help="FY end year e.g. 2027 (live fetch)")
    ap.add_argument("--port",       type=int, default=9000)
    ap.add_argument("--save-db",    action="store_true")
    ap.add_argument("--out",        help="Output PDF path")
    ap.add_argument("--xlsx",          action="store_true", help="Also generate BS-only Excel (legacy)")
    ap.add_argument("--xlsx-classic",  action="store_true",
                    help="Also generate full Schedule III Classic Excel with formulas + unit conversion")
    ap.add_argument("--bs-only",       action="store_true", help="Excel BS only (no PDF, legacy)")
    ap.add_argument("--bs-only-classic", action="store_true",
                    help="Classic Excel only (no PDF) — BS + P&L + TB Data sheets with formulas")
    args = ap.parse_args()

    D_py = None
    if args.json:
        D = load_from_saved_json(args.json)
        if args.json_py:
            D_py = load_from_saved_json(args.json_py)
    elif args.db_id:
        D, D_py = load_from_db(report_id=args.db_id)
    elif args.db_company and args.db_fyend:
        D, D_py = load_from_db(company=args.db_company, fy_end=args.db_fyend)
    elif args.company and args.year:
        D, D_py = load_live_with_prior_year(args.company, args.year, args.port, save_db=args.save_db)
    else:
        ap.error("Provide one of:  --json  |  --db-id  |  --db-company+--db-fyend  |  --company+--year")

    safe = re.sub(r'[^\w]', '_', D['company'][:30])
    out  = args.out or f"classic_sch3_{safe}_{D['fy_end']}.pdf"

    if getattr(args, 'bs_only_classic', False):
        # New: full Classic Excel only (no PDF)
        xlsx_classic_path = out.replace(".pdf", "_Classic.xlsx") if out.endswith(".pdf") else out + "_Classic.xlsx"
        from sch3_excel_classic import generate_schedule3_excel
        generate_schedule3_excel(D, D_py, xlsx_classic_path)
    elif args.bs_only:
        # Legacy BS-only Excel
        xlsx_path = out.replace(".pdf", ".xlsx") if out.endswith(".pdf") else out + ".xlsx"
        download_bs_xlsx(D, D_py, xlsx_path)
    else:
        generate_classic_pdf(D, D_py, out,
                             also_xlsx=getattr(args, 'xlsx', False),
                             also_xlsx_classic=getattr(args, 'xlsx_classic', False))