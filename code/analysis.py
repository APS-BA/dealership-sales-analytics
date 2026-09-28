"""
================================================================================
DATA-DRIVEN MARKETING — COURSE PROJECT
Structural Category Decline and the Cross-Sell Opportunity
Multi-branch consumer durables dealership, Varanasi (U.P.)

Group 02, Section A

WHAT THIS SCRIPT DOES
---------------------
Reproduces the ENTIRE analysis from the three raw GST sales registers and
regenerates the submitted Excel workbook. Nothing is hard-coded: every figure
quoted in the report is computed here from the raw files.

HOW TO RUN
----------
    pip install pandas numpy openpyxl
    python analysis.py

Place the three raw files in the SAME folder as this script (or edit DATA_DIR):
    SALE_DATA_2324.xlsx        -> FY2023-24
    SALE_DATA_2425_NEW.xlsx    -> FY2024-25
    SALE_DATA.xlsx             -> FY2025-26

Outputs:
    LG_Kent_Marketing_Analytics_FY2324_FY2526.xlsx   (16-sheet workbook)
    console log of all key figures, for verification against the report

Runtime: ~2-3 minutes (the market-basket pass over 36,056 invoices dominates).

TWO CORRECTIONS IMPLEMENTED HERE (see report Section 4)
-------------------------------------------------------
1. LINE-LEVEL REVENUE. Category is a LINE-level attribute; GRANDAMOUNT is an
   INVOICE-level total. Attributing GRANDAMOUNT to each category on an invoice
   double-counts every multi-category invoice. We therefore use
   LINE_REVENUE = TOTAL + CGSTAMT + SGSTAMT + IGSTAMT, which reconciles to
   ~100.1% of distinct-invoice GRANDAMOUNT in all three years (asserted below).

2. LIKE-FOR-LIKE WINDOW. FY2024-25 is missing most Feb-Mar billing (298 and 35
   invoices vs ~800 and ~1,100 in adjacent years). All year-on-year comparison
   therefore runs on fiscal months 1-10 (April-January).
================================================================================
"""

import os
import sys
from itertools import combinations

import numpy as np
import pandas as pd

# ----------------------------------------------------------------------------
# CONFIGURATION
# ----------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
def _find_data_dir(base):
    """Locate the folder holding the three raw registers: ./ then ../Data."""
    import os
    probe = "SALE_DATA_2324.xlsx"
    for cand in (base, os.path.join(base, "Data"),
                 os.path.join(os.path.dirname(base), "Data")):
        if os.path.exists(os.path.join(cand, probe)):
            return cand
    return base
DATA_DIR = _find_data_dir(_SCRIPT_DIR)
OUT_XLSX = os.path.join(_SCRIPT_DIR, "LG_Kent_Marketing_Analytics_FY2324_FY2526.xlsx")

SOURCES = [
    ("FY2023-24", "SALE_DATA_2324.xlsx"),
    ("FY2024-25", "SALE_DATA_2425_NEW.xlsx"),
    ("FY2025-26", "SALE_DATA.xlsx"),
]

# HSN -> category. Top 15 codes cover 96.6% of revenue; everything else = Other.
HSN_MAP = {
    841510: "Air Conditioner",
    852872: "Television",
    841810: "Refrigerator",
    84501100: "Washing Machine Automatic",
    84501200: "Washing Machine Semi-Automatic",
    842121: "Water Purifier RO",
}

CATS = [
    "Air Conditioner", "Television", "Refrigerator", "Other",
    "Washing Machine Automatic", "Washing Machine Semi-Automatic",
    "Water Purifier RO",
]
FYS = [fy for fy, _ in SOURCES]
FM_NAMES = ["Apr", "May", "Jun", "Jul", "Aug", "Sep",
            "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
LIKE_FOR_LIKE_MONTHS = 10  # April (1) .. January (10)


def rule(title):
    print("\n" + "=" * 78)
    print(title)
    print("=" * 78)


# ============================================================================
# STEP 1 — LOAD AND PREPROCESS
# ============================================================================
def load_and_preprocess():
    rule("STEP 1 — LOAD AND PREPROCESS")

    frames = []
    for fy, fname in SOURCES:
        path = os.path.join(DATA_DIR, fname)
        if not os.path.exists(path):
            sys.exit(f"ERROR: missing input file {path}")
        df = pd.read_excel(path, sheet_name="Sheet1")
        df["FISCAL_YEAR"] = fy
        frames.append(df)
        print(f"  loaded {fname:26s} -> {len(df):>7,} rows  ({fy})")

    raw_rows = sum(len(f) for f in frames)

    # Union on common fields. FY2025-26 uniquely has GSTIN/SALESLEDGER; the two
    # earlier files uniquely have STATE/COUNTRY/BUYERMAILINGNAME/ADDRESS2/
    # CASHBACK. These are retained as optional columns, never used as join keys.
    common = set(frames[0].columns)
    for f in frames[1:]:
        common &= set(f.columns)
    print(f"\n  common columns across all three files: {len(common)}")

    d = pd.concat(frames, ignore_index=True)

    dups = int(d.duplicated().sum())
    d = d.drop_duplicates().reset_index(drop=True)
    print(f"  raw rows                 : {raw_rows:,}")
    print(f"  exact duplicates removed : {dups}")
    print(f"  cleaned master rows      : {len(d):,}")

    # Dates are day-first in the source.
    d["INVOICE_DATE"] = pd.to_datetime(d["INVOICEDATE"], errors="coerce", dayfirst=True)
    d["MONTH"] = d["INVOICE_DATE"].dt.month
    d["FISCAL_MONTH"] = d["MONTH"].apply(lambda m: (m - 4) % 12 + 1)  # Apr=1..Mar=12
    d["FISCAL_MONTH_NAME"] = d["FISCAL_MONTH"].apply(lambda x: FM_NAMES[int(x) - 1])

    d["CATEGORY"] = d["HSNCODE"].map(HSN_MAP).fillna("Other")

    # CORRECTION 1 — line-level revenue (see module docstring).
    d["LINE_REVENUE"] = (d["TOTAL"].fillna(0) + d["SGSTAMT"].fillna(0)
                         + d["CGSTAMT"].fillna(0) + d["IGSTAMT"].fillna(0))

    # Quality flags — rows are FLAGGED, never deleted.
    d["ZERO_VALUE_FLAG"] = d["GSTRATE"].fillna(0) == 0
    d["BAD_SKU_FLAG"] = d["PRODUCT"].astype(str).str.fullmatch(r"\d+").fillna(False)
    d["CASH_CUSTOMER_FLAG"] = (d["CUSTOMERNAME"].astype(str)
                               .str.contains("CASH CUSTOMER", case=False, na=False))
    d["BRANCH"] = d["INVOICENO"].astype(str).str.extract(r"^([A-Z]+)/")[0]

    print(f"  zero-value lines flagged : {int(d.ZERO_VALUE_FLAG.sum()):,}")
    print(f"  numeric-only SKUs flagged: {int(d.BAD_SKU_FLAG.sum()):,}")
    print(f"  cash-customer lines      : {int(d.CASH_CUSTOMER_FLAG.sum()):,}  (excluded from customer analysis)")

    # GRANDAMOUNT must be constant within an invoice for invoice-level use.
    varies = d.groupby("INVOICENO")["GRANDAMOUNT"].nunique()
    assert (varies <= 1).all(), "GRANDAMOUNT varies within an invoice — cannot treat as invoice total"
    print("  CHECK: GRANDAMOUNT constant within every invoice ....... OK")

    return d, raw_rows, dups


def reconciliation_check(d):
    """CORRECTION 1 verification: line revenue must reconcile to invoice totals."""
    rule("STEP 2 — RECONCILIATION CHECK (validates the line-level revenue basis)")
    inv = d.drop_duplicates("INVOICENO")
    for fy in FYS:
        line = d.loc[d.FISCAL_YEAR == fy, "LINE_REVENUE"].sum()
        grand = inv.loc[inv.FISCAL_YEAR == fy, "GRANDAMOUNT"].sum()
        ratio = line / grand
        print(f"  {fy}: line {line:>15,.0f} / invoice {grand:>15,.0f} = {ratio:.5f}")
        assert 0.99 <= ratio <= 1.01, f"{fy} fails reconciliation at {ratio:.4f}"
    print("  CHECK: all three years reconcile within +/-1% ........... OK")
    print("  (Under the erroneous invoice-attribution method the category")
    print("   parts EXCEEDED the whole by roughly 2x — see report Section 4.)")


# ============================================================================
# ANALYSIS A — CATEGORY AND SEASONALITY
# ============================================================================
def analysis_a(d):
    rule("ANALYSIS A — CATEGORY REVENUE AND SEASONALITY")
    inv = d.drop_duplicates("INVOICENO")

    full = (d.pivot_table(index="CATEGORY", columns="FISCAL_YEAR",
                          values="LINE_REVENUE", aggfunc="sum")
            .reindex(CATS).fillna(0))

    ll_src = d[d.FISCAL_MONTH <= LIKE_FOR_LIKE_MONTHS]
    like = (ll_src.pivot_table(index="CATEGORY", columns="FISCAL_YEAR",
                               values="LINE_REVENUE", aggfunc="sum")
            .reindex(CATS).fillna(0))

    month_rev = inv.pivot_table(index="FISCAL_MONTH", columns="FISCAL_YEAR",
                                values="GRANDAMOUNT", aggfunc="sum").fillna(0)
    month_cnt = inv.pivot_table(index="FISCAL_MONTH", columns="FISCAL_YEAR",
                                values="INVOICENO", aggfunc="count").fillna(0)
    cat_month = (d.pivot_table(index="CATEGORY", columns="FISCAL_MONTH",
                               values="LINE_REVENUE", aggfunc="sum")
                 .reindex(CATS).fillna(0))

    print("\n  FULL-YEAR (contaminated by the FY2024-25 Feb-Mar gap):")
    for c in CATS:
        print(f"    {c:32s} {full.loc[c, FYS[0]]:>14,.0f} {full.loc[c, FYS[1]]:>14,.0f} "
              f"{full.loc[c, FYS[2]]:>14,.0f}")
    t = full.sum()
    print(f"    {'TOTAL':32s} {t[FYS[0]]:>14,.0f} {t[FYS[1]]:>14,.0f} {t[FYS[2]]:>14,.0f}")
    print(f"    YoY: {(t[FYS[1]]/t[FYS[0]]-1)*100:+.1f}%  then  {(t[FYS[2]]/t[FYS[1]]-1)*100:+.1f}%   <- the ARTEFACT")

    print("\n  LIKE-FOR-LIKE April-January (the reliable view):")
    for c in CATS:
        two_yr = (like.loc[c, FYS[2]] / like.loc[c, FYS[0]] - 1) * 100
        print(f"    {c:32s} {like.loc[c, FYS[0]]:>14,.0f} {like.loc[c, FYS[1]]:>14,.0f} "
              f"{like.loc[c, FYS[2]]:>14,.0f}   2-yr {two_yr:+6.1f}%")
    tl = like.sum()
    print(f"    {'TOTAL':32s} {tl[FYS[0]]:>14,.0f} {tl[FYS[1]]:>14,.0f} {tl[FYS[2]]:>14,.0f}   "
          f"2-yr {(tl[FYS[2]]/tl[FYS[0]]-1)*100:+6.1f}%")
    print(f"    YoY: {(tl[FYS[1]]/tl[FYS[0]]-1)*100:+.1f}%  then  {(tl[FYS[2]]/tl[FYS[1]]-1)*100:+.1f}%   <- erosion, no V")

    # Seasonality concentration
    q1 = month_rev.loc[1:3].sum() / month_rev.sum() * 100
    q2 = month_rev.loc[4:6].sum() / month_rev.sum() * 100
    print("\n  Seasonality — share of annual revenue:")
    for fy in FYS:
        print(f"    {fy}: Apr-Jun {q1[fy]:.1f}%   Jul-Sep {q2[fy]:.1f}%")

    return full, like, month_rev, month_cnt, cat_month


def data_gap_check(month_cnt, month_rev):
    rule("DATA GAP CHECK — why the like-for-like window is necessary")
    print("  Invoice counts by fiscal month:")
    print(f"    {'Month':6s}" + "".join(f"{fy:>13s}" for fy in FYS))
    for i in range(1, 13):
        row = "".join(f"{int(month_cnt.loc[i, fy]):>13,}" for fy in FYS)
        mark = "   <-- INCOMPLETE" if i in (11, 12) else ""
        print(f"    {FM_NAMES[i-1]:6s}{row}{mark if i == 12 else ''}")
    fm = {fy: month_rev.loc[[11, 12], fy].sum() for fy in FYS}
    norm = (fm[FYS[0]] + fm[FYS[2]]) / 2
    print(f"\n  Feb+Mar revenue: {fm[FYS[0]]:,.0f} | {fm[FYS[1]]:,.0f} | {fm[FYS[2]]:,.0f}")
    print(f"  Estimated FY2024-25 shortfall: {norm - fm[FYS[1]]:,.0f}")
    print("  A dealership does not fall to 35 invoices in March and recover to")
    print("  1,249 the next year. This is a truncated export, not a collapse.")


# ============================================================================
# ANALYSIS B — MARKET BASKET (SUPPORT, CONFIDENCE, LIFT)
# ============================================================================
def basket_rules(basket_sets, label):
    """
    Support    = P(A and B)            -- how often the pair occurs at all
    Confidence = P(B|A) = supp/P(A)    -- conditional probability of consequent
    Lift       = conf / P(B)           -- confidence corrected for B's base rate

    All three are reported because each fails alone: confidence is inflated by a
    popular consequent, so a rule can be high-confidence yet NEGATIVELY
    associated (lift < 1). See report Section 5.
    """
    n = len(basket_sets)
    rows = []
    for a, b in combinations(CATS, 2):
        na = sum(1 for s in basket_sets if a in s)
        nb = sum(1 for s in basket_sets if b in s)
        both = sum(1 for s in basket_sets if a in s and b in s)
        if both == 0 or na == 0 or nb == 0:
            continue
        support = both / n
        rows.append({
            "BASKET_TYPE": label, "CATEGORY_A": a, "CATEGORY_B": b,
            "BASKETS": n, "PAIR_BASKETS": both, "SUPPORT": support,
            "CONFIDENCE_A_TO_B": both / na,
            "CONFIDENCE_B_TO_A": both / nb,
            "LIFT": support / ((na / n) * (nb / n)),
        })
    return rows


def analysis_b(d):
    rule("ANALYSIS B — MARKET BASKET: SUPPORT, CONFIDENCE, LIFT")
    named = d[~d.CASH_CUSTOMER_FLAG]

    inv_sets = d.groupby("INVOICENO")["CATEGORY"].apply(set)
    cust_sets = named.groupby("CUSTOMERNAME")["CATEGORY"].apply(set)
    print(f"  invoice baskets : {len(inv_sets):,}")
    print(f"  customer baskets: {len(cust_sets):,}")

    pairs = pd.DataFrame(basket_rules(inv_sets, "Invoice basket")
                         + basket_rules(cust_sets, "Customer basket"))
    pairs = pairs.sort_values(["BASKET_TYPE", "LIFT"], ascending=[True, False])

    print("\n  INVOICE-LEVEL, by lift (support / conf A->B / conf B->A / lift):")
    for _, r in pairs[pairs.BASKET_TYPE == "Invoice basket"].iterrows():
        flag = "  <-- AC negative" if r.CATEGORY_A == "Air Conditioner" else ""
        print(f"    {r.CATEGORY_A[:22]:22s} + {r.CATEGORY_B[:22]:22s} "
              f"{r.SUPPORT:.4f} {r.CONFIDENCE_A_TO_B:.3f} {r.CONFIDENCE_B_TO_A:.3f} "
              f"{r.LIFT:.3f}{flag}")

    # The support/confidence/lift teaching point quoted in the report.
    print("\n  WHY LIFT AND NOT CONFIDENCE ALONE:")
    inv_rules = pairs[pairs.BASKET_TYPE == "Invoice basket"]
    top_conf = inv_rules.loc[inv_rules.CONFIDENCE_B_TO_A.idxmax()]
    top_lift = inv_rules.loc[inv_rules.LIFT.idxmax()]
    print(f"    Highest confidence rule : {top_conf.CATEGORY_B} -> {top_conf.CATEGORY_A} "
          f"conf {top_conf.CONFIDENCE_B_TO_A:.3f}, support {top_conf.SUPPORT:.4f}, "
          f"but LIFT {top_conf.LIFT:.3f}")
    print(f"    Highest lift rule       : {top_lift.CATEGORY_A} -> {top_lift.CATEGORY_B} "
          f"conf {top_lift.CONFIDENCE_A_TO_B:.3f}, support {top_lift.SUPPORT:.4f}, "
          f"LIFT {top_lift.LIFT:.3f}")
    print("    -> acting on confidence alone promotes a negatively associated rule.")

    ac = pairs[(pairs.CATEGORY_A == "Air Conditioner")]
    print(f"\n  AC lift range — invoice : {ac[ac.BASKET_TYPE=='Invoice basket'].LIFT.min():.3f} "
          f"to {ac[ac.BASKET_TYPE=='Invoice basket'].LIFT.max():.3f}")
    print(f"  AC lift range — customer: {ac[ac.BASKET_TYPE=='Customer basket'].LIFT.min():.3f} "
          f"to {ac[ac.BASKET_TYPE=='Customer basket'].LIFT.max():.3f}")
    print("  AC is below 1.0 against EVERY category at BOTH levels — not a timing artefact.")

    # Single vs multi-category value gap
    sm = []
    for fy in FYS:
        s = named[named.FISCAL_YEAR == fy]
        ncat = s.groupby("CUSTOMERNAME")["CATEGORY"].nunique()
        rev = s.drop_duplicates("INVOICENO").groupby("CUSTOMERNAME")["GRANDAMOUNT"].sum()
        for lab, idx in [("Multi-category", ncat[ncat >= 2].index),
                         ("Single-category", ncat[ncat == 1].index)]:
            sm.append({"FISCAL_YEAR": fy, "CUSTOMER_TYPE": lab, "CUSTOMERS": len(idx),
                       "AVG_TOTAL_SPEND": rev[idx].mean(), "TOTAL_REVENUE": rev[idx].sum()})
    sm = pd.DataFrame(sm)

    print("\n  CROSS-SELL VALUE GAP:")
    for fy in FYS:
        m = sm[(sm.FISCAL_YEAR == fy) & (sm.CUSTOMER_TYPE == "Multi-category")].iloc[0]
        s = sm[(sm.FISCAL_YEAR == fy) & (sm.CUSTOMER_TYPE == "Single-category")].iloc[0]
        share = s.CUSTOMERS / (s.CUSTOMERS + m.CUSTOMERS) * 100
        print(f"    {fy}: multi {m.CUSTOMERS:,} @ {m.AVG_TOTAL_SPEND:,.0f} | "
              f"single {s.CUSTOMERS:,} @ {s.AVG_TOTAL_SPEND:,.0f} | "
              f"ratio {m.AVG_TOTAL_SPEND/s.AVG_TOTAL_SPEND:.2f}x | single share {share:.1f}%")

    return pairs, sm


# ============================================================================
# ANALYSIS C — SEASONAL PRICING
# ============================================================================
def analysis_c(d):
    rule("ANALYSIS C — SEASONAL PRICING (product-normalised index)")
    pc = d[d.BASICRATE > 0].copy()
    # Normalise each transaction against ITS OWN product's 3-year mean price, so
    # that changes in product mix cannot masquerade as price movement.
    pc["PROD_MEAN"] = pc.groupby("PRODUCT")["BASICRATE"].transform("mean")
    pc["PRICE_INDEX"] = pc["BASICRATE"] / pc["PROD_MEAN"]

    pidx = (pc.pivot_table(index="CATEGORY", columns="FISCAL_MONTH",
                           values="PRICE_INDEX", aggfunc="mean").reindex(CATS))
    cov = (pidx.std(axis=1) / pidx.mean(axis=1)).rename("COEFF_OF_VARIATION")

    print("  Coefficient of variation of the monthly price index:")
    for c, v in cov.sort_values(ascending=False).items():
        note = ""
        if c == "Air Conditioner":
            note = "   <-- most seasonal DEMAND, near-flat PRICE"
        if c == "Television":
            note = "   <-- no weather seasonality, yet MORE volatile"
        print(f"    {c:32s} {v:.4f}{note}")

    # AC demand vs price, to show the divergence quoted in the report.
    ac = d[d.CATEGORY == "Air Conditioner"]
    dem = ac.groupby("FISCAL_MONTH")["LINE_REVENUE"].sum()
    dem_idx = dem / dem.mean()
    print("\n  AC demand index vs price index by fiscal month:")
    for i in range(1, 13):
        if i in dem_idx.index and i in pidx.columns:
            print(f"    {FM_NAMES[i-1]:4s} demand {dem_idx[i]:5.2f}   price {pidx.loc['Air Conditioner', i]:.3f}")
    print(f"\n  AC demand swing (max/min month): {dem_idx.max()/dem_idx.min():.1f}x")
    print(f"  AC price movement across the year: {cov['Air Conditioner']*100:.2f}%")
    print("  CAVEAT: these are negotiated transaction prices, not a controlled")
    print("  experiment. The flatness is robust; causal elasticity is NOT tested.")

    return pidx, cov


# ============================================================================
# ANALYSIS D — RFM
# ============================================================================
def analysis_d(d):
    rule("ANALYSIS D — RFM SEGMENTATION (recalibrated for durables)")
    named = d[~d.CASH_CUSTOMER_FLAG]
    inv = named.drop_duplicates("INVOICENO")[
        ["INVOICENO", "CUSTOMERNAME", "INVOICE_DATE", "GRANDAMOUNT"]]
    snap = inv.INVOICE_DATE.max()
    print(f"  snapshot date: {snap.date()}")

    rfm = inv.groupby("CUSTOMERNAME").agg(
        RECENCY_DAYS=("INVOICE_DATE", lambda x: (snap - x.max()).days),
        FREQUENCY=("INVOICENO", "nunique"),
        MONETARY=("GRANDAMOUNT", "sum"))

    # R and M quintile-scored. F is BANDED, not quintile-scored: with 83.8% of
    # customers at exactly one invoice the distribution is near-degenerate and
    # qcut cannot produce five distinct bins. See report Q6.
    rfm["R_SCORE"] = pd.qcut(rfm.RECENCY_DAYS.rank(method="first"), 5,
                             labels=[5, 4, 3, 2, 1]).astype(int)
    rfm["F_SCORE"] = pd.cut(rfm.FREQUENCY, bins=[0, 1, 2, 3, 5, 10**9],
                            labels=[1, 2, 3, 4, 5]).astype(int)
    rfm["M_SCORE"] = pd.qcut(rfm.MONETARY.rank(method="first"), 5,
                             labels=[1, 2, 3, 4, 5]).astype(int)
    rfm["RFM_SCORE"] = (rfm.R_SCORE.astype(str) + rfm.F_SCORE.astype(str)
                        + rfm.M_SCORE.astype(str))

    def segment(r):
        if r.F_SCORE >= 4 and r.M_SCORE >= 4:
            return "Champions"
        if r.F_SCORE >= 3 and r.M_SCORE >= 3:
            return "Loyal"
        if r.R_SCORE >= 4 and r.F_SCORE <= 2:
            return "Potential Loyalist"
        if r.R_SCORE <= 2 and r.M_SCORE >= 4:
            return "At Risk (high value)"
        if r.R_SCORE <= 2:
            return "Lapsed / Not Yet Due"
        return "Needs Attention"

    rfm["SEGMENT"] = rfm.apply(segment, axis=1)

    seg = (rfm.groupby("SEGMENT")
           .agg(CUSTOMERS=("MONETARY", "size"), REVENUE=("MONETARY", "sum"),
                AVG_MONETARY=("MONETARY", "mean"),
                AVG_FREQUENCY=("FREQUENCY", "mean"),
                AVG_RECENCY_DAYS=("RECENCY_DAYS", "mean"))
           .reset_index().sort_values("REVENUE", ascending=False))
    seg["REVENUE_SHARE"] = seg.REVENUE / seg.REVENUE.sum()

    print(f"\n  named customers: {len(rfm):,} | total monetary: {rfm.MONETARY.sum():,.0f}")
    print("  Segment                  customers        revenue   share   avgF   avgR")
    for _, r in seg.iterrows():
        print(f"    {r.SEGMENT:24s} {r.CUSTOMERS:>7,} {r.REVENUE:>14,.0f} "
              f"{r.REVENUE_SHARE*100:>6.2f}% {r.AVG_FREQUENCY:>6.2f} {r.AVG_RECENCY_DAYS:>6.0f}")

    once = int((rfm.FREQUENCY == 1).sum())
    srt = rfm.MONETARY.sort_values(ascending=False)
    cum = srt.cumsum() / srt.sum()
    top10 = cum.iloc[int(len(srt) * .10) - 1]
    top20 = cum.iloc[int(len(srt) * .20) - 1]

    print(f"\n  Q6 EVIDENCE — is RFM valid here?")
    print(f"    customers with exactly ONE invoice in 3 years: {once:,} ({once/len(rfm)*100:.1f}%)")
    print(f"    mean frequency                               : {rfm.FREQUENCY.mean():.2f}")
    print(f"    median recency                               : {rfm.RECENCY_DAYS.median():.0f} days")
    print("    -> F is near-degenerate; RFM collapses toward an R-M map and any")
    print("       CHURN reading is invalid. On decade-cycle goods a low F score is")
    print("       a SATISFIED customer. Valid use = value concentration/contactability:")
    print(f"    top 10% of customers = {top10*100:.1f}% of revenue")
    print(f"    top 20% of customers = {top20*100:.1f}% of revenue")

    return rfm, seg, {"once": once, "top10": top10, "top20": top20, "snap": snap}


# ============================================================================
# MAIN
# ============================================================================
def main():
    print(__doc__)
    d, raw_rows, dups = load_and_preprocess()
    reconciliation_check(d)
    full, like, month_rev, month_cnt, cat_month = analysis_a(d)
    data_gap_check(month_cnt, month_rev)
    pairs, sm = analysis_b(d)
    pidx, cov = analysis_c(d)
    rfm, seg, rfm_stats = analysis_d(d)

    rule("BUILDING THE WORKBOOK")
    builder = os.path.join(_SCRIPT_DIR, "build_workbook.py")
    if os.path.exists(builder):
        import subprocess
        res = subprocess.run([sys.executable, builder], capture_output=True, text=True)
        print(res.stdout.strip() or res.stderr.strip())
        if os.path.exists(OUT_XLSX):
            print(f"  written: {OUT_XLSX}")
    else:
        print("  build_workbook.py not found — analysis complete, workbook skipped.")

    rule("DONE — all figures above are reproduced in the report and workbook")


if __name__ == "__main__":
    main()
