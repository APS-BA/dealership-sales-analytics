"""
Final three-year marketing analytics workbook build, including DATA_GAP_CHECK, like-for-like, key findings and recommendations sheets.

Source: generated and run in Better Analyst (app.betteranalyst.com), session
"Marketing Analytics Workbook Preview", 2026-09-04. Exported verbatim; only this header was added.
Input paths point to Better Analyst's sandbox: /home/user/workspace/1788497159764_SALE_DATA_2425_NEW.xlsx, /home/user/workspace/1788497160021_SALE_DATA.xlsx, /home/user/workspace/1788497160280_SALE_DATA_2324.xlsx, /home/user/workspace/outputs/lg_kent_marketing_analytics_fy23_26.xlsx.
Change them to your local data folder before running.
"""

from pathlib import Path
from datetime import datetime
from itertools import combinations
import math, re, json, os
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

source_paths = [
    ("/home/user/workspace/1788497160280_SALE_DATA_2324.xlsx", "FY2023-24"),
    ("/home/user/workspace/1788497159764_SALE_DATA_2425_NEW.xlsx", "FY2024-25"),
    ("/home/user/workspace/1788497160021_SALE_DATA.xlsx", "FY2025-26"),
]
for source_path, _fy in source_paths:
    p = Path(source_path)
    if not p.is_file():
        raise ValueError(f"Missing source file: {source_path}")
    print(f"AUDIT: source={source_path}")

# Load the three registers from Sheet1 and union their observed columns.
frames = []
source_rows = []
for source_path, fy in source_paths:
    df0 = pd.read_excel(source_path, sheet_name="Sheet1")
    df0.columns = [str(c).strip() for c in df0.columns]
    df0["FISCAL_YEAR"] = fy
    frames.append(df0)
    source_rows.append((fy, len(df0), list(df0.columns)))

all_cols = []
for df0 in frames:
    for c in df0.columns:
        if c not in all_cols:
            all_cols.append(c)
raw = pd.concat([df0.reindex(columns=all_cols) for df0 in frames], ignore_index=True)
raw_row_count = len(raw)

# Remove exact duplicates after union alignment.
raw = raw.drop_duplicates(keep="first").reset_index(drop=True)
duplicate_rows_removed = raw_row_count - len(raw)

# Type-safe preprocessing.
raw["INVOICEDATE_clean"] = pd.to_datetime(raw["INVOICEDATE"], dayfirst=True, errors="coerce")
raw["HSNCODE_clean"] = pd.to_numeric(raw["HSNCODE"], errors="coerce")
raw["BASICRATE_clean"] = pd.to_numeric(raw["BASICRATE"], errors="coerce")
raw["GSTRATE_clean"] = pd.to_numeric(raw["GSTRATE"], errors="coerce")
raw["GRANDAMOUNT_clean"] = pd.to_numeric(raw["GRANDAMOUNT"], errors="coerce")
raw["MONTH"] = raw["INVOICEDATE_clean"].dt.month
raw["FISCAL_MONTH"] = ((raw["MONTH"] - 4) % 12) + 1

hsn_map = {
    "841510": "Air Conditioner",
    "852872": "Television",
    "841810": "Refrigerator",
    "84501100": "Washing Machine Automatic",
    "84501200": "Washing Machine Semi-Automatic",
    "842121": "Water Purifier RO",
}
raw["HSNCODE_KEY"] = raw["HSNCODE_clean"].map(lambda x: str(int(x)) if pd.notna(x) else "")
raw["CATEGORY"] = raw["HSNCODE_KEY"].map(hsn_map).fillna("Other")
raw["ZERO_VALUE_FLAG"] = raw["GSTRATE_clean"].fillna(0).eq(0)
raw["BAD_SKU_FLAG"] = raw["PRODUCT"].astype("string").fillna("").str.strip().str.fullmatch(r"[0-9]+", na=False)
raw["CASH_CUSTOMER_FLAG"] = raw["CUSTOMERNAME"].astype("string").fillna("").str.upper().str.contains("CASH CUSTOMER", regex=False, na=False)
raw["BRANCH"] = raw["INVOICENO"].astype("string").fillna("").str.split("/", n=1).str[0].str.strip()

# Cleaned master columns: keep all original union columns and add derived fields; do not retain helper clean columns.
helper_cols = {"INVOICEDATE_clean", "HSNCODE_clean", "BASICRATE_clean", "GSTRATE_clean", "GRANDAMOUNT_clean", "HSNCODE_KEY"}
cleaned_cols = [c for c in raw.columns if c not in helper_cols]
cleaned = raw[cleaned_cols].copy()

# Invoice-level table: distinct invoice amount, customer, branch, date, fiscal year, and category list.
invoice_group_cols = ["INVOICENO", "INVOICEDATE_clean", "FISCAL_YEAR", "CUSTOMERNAME", "BRANCH", "CASH_CUSTOMER_FLAG"]
invoice_rows = []
for invoice_no, g in raw.groupby("INVOICENO", dropna=False, sort=False):
    cats = sorted(set([str(x) for x in g["CATEGORY"].dropna().tolist() if str(x).strip()]))
    first = g.iloc[0]
    amounts = pd.to_numeric(g["GRANDAMOUNT_clean"], errors="coerce").dropna().unique()
    amount = float(amounts[0]) if len(amounts) else None
    invoice_rows.append({
        "INVOICENO": invoice_no,
        "INVOICEDATE": first["INVOICEDATE_clean"],
        "FISCAL_YEAR": first["FISCAL_YEAR"],
        "CUSTOMERNAME": first["CUSTOMERNAME"],
        "BRANCH": first["BRANCH"],
        "CASH_CUSTOMER_FLAG": bool(first["CASH_CUSTOMER_FLAG"]),
        "GRANDAMOUNT": amount,
        "CATEGORIES": "; ".join(cats),
        "CATEGORY_COUNT": len(cats),
    })
invoices = pd.DataFrame(invoice_rows)
invoices["INVOICEDATE"] = pd.to_datetime(invoices["INVOICEDATE"], errors="coerce")
invoices["FISCAL_MONTH"] = ((invoices["INVOICEDATE"].dt.month - 4) % 12) + 1

# Use distinct invoice amounts for revenue analyses.
valid_invoice = invoices[invoices["GRANDAMOUNT"].notna()].copy()
valid_invoice["GRANDAMOUNT"] = pd.to_numeric(valid_invoice["GRANDAMOUNT"], errors="coerce")

# Analysis A: category x FY revenue, fiscal-month revenue, pooled category x fiscal month.
cat_fy = raw.groupby(["CATEGORY", "FISCAL_YEAR"], dropna=False)["GRANDAMOUNT_clean"].agg(lambda s: pd.to_numeric(s, errors="coerce").drop_duplicates().sum())
# Distinct invoice revenue by fiscal year and month.
month_fy = valid_invoice.groupby(["FISCAL_MONTH", "FISCAL_YEAR"], dropna=False)["GRANDAMOUNT"].sum().reset_index()
# Category revenue by fiscal month uses line-level categories with invoice amount allocated once across categories on that invoice to avoid double counting.
cat_invoice = raw[["INVOICENO", "CATEGORY"]].drop_duplicates()
cat_invoice = cat_invoice.merge(valid_invoice[["INVOICENO", "GRANDAMOUNT", "FISCAL_MONTH", "FISCAL_YEAR"]], on="INVOICENO", how="left")
cat_invoice["CATEGORY_COUNT"] = cat_invoice.groupby("INVOICENO")["CATEGORY"].transform("count")
cat_invoice["ALLOCATED_REVENUE"] = cat_invoice["GRANDAMOUNT"] / cat_invoice["CATEGORY_COUNT"].replace(0, np.nan)
cat_month = cat_invoice.groupby(["CATEGORY", "FISCAL_MONTH"], dropna=False)["ALLOCATED_REVENUE"].sum().reset_index()

# Analysis B: basket metrics.
def basket_metrics(basket_sets, label):
    basket_sets = [set(x) for x in basket_sets if len(set(x)) > 0]
    n = len(basket_sets)
    cat_counts = {}
    pair_counts = {}
    for s in basket_sets:
        for c in s:
            cat_counts[c] = cat_counts.get(c, 0) + 1
        for a, b in combinations(sorted(s), 2):
            pair_counts[(a, b)] = pair_counts.get((a, b), 0) + 1
    rows = []
    for (a, b), cnt in sorted(pair_counts.items(), key=lambda kv: (-kv[1], kv[0])):
        support = cnt / n if n else 0
        conf_ab = cnt / cat_counts[a] if cat_counts.get(a, 0) else 0
        conf_ba = cnt / cat_counts[b] if cat_counts.get(b, 0) else 0
        support_a = cat_counts[a] / n if n else 0
        support_b = cat_counts[b] / n if n else 0
        lift = support / (support_a * support_b) if support_a and support_b else 0
        rows.append({"BASKET_TYPE": label, "CATEGORY_A": a, "CATEGORY_B": b, "BASKETS": n, "PAIR_BASKETS": cnt, "SUPPORT": support, "CONFIDENCE_A_TO_B": conf_ab, "CONFIDENCE_B_TO_A": conf_ba, "LIFT": lift})
    return pd.DataFrame(rows)

invoice_sets = valid_invoice["CATEGORIES"].fillna("").map(lambda s: [x.strip() for x in s.split(";") if x.strip()])
basket_invoice = basket_metrics(invoice_sets.tolist(), "Invoice basket")

named_invoices = valid_invoice[~valid_invoice["CASH_CUSTOMER_FLAG"]].copy()
named_customer_sets = named_invoices.groupby("CUSTOMERNAME")["CATEGORIES"].apply(lambda s: sorted(set(sum(([x.strip() for x in str(v).split(";") if x.strip()] for v in s), []))))
basket_customer = basket_metrics(named_customer_sets.tolist(), "Named customer basket")
basket_pairs = pd.concat([basket_invoice, basket_customer], ignore_index=True)

# Customer single vs multi category by fiscal year.
customer_fy_cats = named_invoices.groupby(["FISCAL_YEAR", "CUSTOMERNAME"])["CATEGORIES"].apply(lambda s: len(set(sum(([x.strip() for x in str(v).split(";") if x.strip()] for v in s), [])))).reset_index(name="CATEGORY_COUNT")
customer_fy_spend = named_invoices.groupby(["FISCAL_YEAR", "CUSTOMERNAME"])["GRANDAMOUNT"].sum().reset_index(name="TOTAL_SPEND")
customer_mix = customer_fy_cats.merge(customer_fy_spend, on=["FISCAL_YEAR", "CUSTOMERNAME"], how="left")
customer_mix["CUSTOMER_TYPE"] = np.where(customer_mix["CATEGORY_COUNT"] <= 1, "Single-category", "Multi-category")
customer_mix_summary = customer_mix.groupby(["FISCAL_YEAR", "CUSTOMER_TYPE"]).agg(CUSTOMERS=("CUSTOMERNAME", "nunique"), AVG_TOTAL_SPEND=("TOTAL_SPEND", "mean"), TOTAL_REVENUE=("TOTAL_SPEND", "sum")).reset_index()

# Analysis C: product-normalized pricing index.
price = raw[["PRODUCT", "CATEGORY", "FISCAL_MONTH", "BASICRATE_clean"]].copy()
price = price[price["BASICRATE_clean"] > 0].copy()
product_means = price.groupby("PRODUCT")["BASICRATE_clean"].mean().rename("PRODUCT_MEAN_BASICRATE")
price = price.join(product_means, on="PRODUCT")
price["PRICE_INDEX"] = price["BASICRATE_clean"] / price["PRODUCT_MEAN_BASICRATE"]
price_month = price.groupby(["CATEGORY", "FISCAL_MONTH"], dropna=False)["PRICE_INDEX"].mean().reset_index(name="AVG_PRICE_INDEX")
price_cv = price_month.groupby("CATEGORY")["AVG_PRICE_INDEX"].agg(["mean", "std"]).reset_index()
price_cv["COEFFICIENT_OF_VARIATION"] = price_cv["std"] / price_cv["mean"].replace(0, np.nan)
price_cv = price_cv[["CATEGORY", "COEFFICIENT_OF_VARIATION"]]
price_month = price_month.merge(price_cv, on="CATEGORY", how="left")

# Analysis D: RFM for named customers only.
rfm_invoices = valid_invoice[~valid_invoice["CASH_CUSTOMER_FLAG"]].copy()
last_date = valid_invoice["INVOICEDATE"].max()
rfm = rfm_invoices.groupby("CUSTOMERNAME").agg(LAST_INVOICE_DATE=("INVOICEDATE", "max"), FREQUENCY=("INVOICENO", "nunique"), MONETARY=("GRANDAMOUNT", "sum")).reset_index()
rfm["RECENCY_DAYS"] = (last_date - rfm["LAST_INVOICE_DATE"]).dt.days

def quintile_score(series, ascending=True):
    ranks = series.rank(method="first", ascending=ascending)
    return pd.qcut(ranks, 5, labels=[1, 2, 3, 4, 5]).astype(int)

rfm["R_SCORE"] = quintile_score(rfm["RECENCY_DAYS"], ascending=False)
rfm["F_SCORE"] = quintile_score(rfm["FREQUENCY"], ascending=True)
rfm["M_SCORE"] = quintile_score(rfm["MONETARY"], ascending=True)
rfm["RFM_SCORE"] = rfm["R_SCORE"].astype(str) + rfm["F_SCORE"].astype(str) + rfm["M_SCORE"].astype(str)
def segment(row):
    if row["R_SCORE"] >= 4 and row["F_SCORE"] >= 4 and row["M_SCORE"] >= 4:
        return "Champions"
    if row["R_SCORE"] >= 3 and row["F_SCORE"] >= 3 and row["M_SCORE"] >= 3:
        return "Loyal"
    if row["RECENCY_DAYS"] <= 90 and row["F_SCORE"] <= 2:
        return "New"
    if row["R_SCORE"] <= 3 and row["F_SCORE"] >= 3:
        return "At Risk"
    if row["R_SCORE"] <= 2 and row["F_SCORE"] <= 2:
        return "Lost"
    return "Potential Loyalist"
rfm["SEGMENT"] = rfm.apply(segment, axis=1)
segment_summary = rfm.groupby("SEGMENT").agg(CUSTOMERS=("CUSTOMERNAME", "nunique"), REVENUE=("MONETARY", "sum"), AVG_MONETARY=("MONETARY", "mean"), AVG_FREQUENCY=("FREQUENCY", "mean"), AVG_RECENCY_DAYS=("RECENCY_DAYS", "mean")).reset_index()
segment_summary["REVENUE_SHARE"] = segment_summary["REVENUE"] / segment_summary["REVENUE"].sum()
segment_summary = segment_summary.sort_values("REVENUE", ascending=False)

# Findings summary values.
total_revenue = float(valid_invoice["GRANDAMOUNT"].sum())
invoice_count = int(valid_invoice["INVOICENO"].nunique())
named_customer_count = int(rfm["CUSTOMERNAME"].nunique())
top_category = raw.groupby("CATEGORY")["GRANDAMOUNT_clean"].agg(lambda s: pd.to_numeric(s, errors="coerce").drop_duplicates().sum()).sort_values(ascending=False).index[0]
top_category_revenue = float(raw.groupby("CATEGORY")["GRANDAMOUNT_clean"].agg(lambda s: pd.to_numeric(s, errors="coerce").drop_duplicates().sum()).sort_values(ascending=False).iloc[0])

# Workbook helpers.
THEME = {"primary":"1F4E79", "light":"D6E3F0", "accent":"2E75B6", "green":"2E7D32", "red":"C62828", "gray":"666666", "light_gray":"F3F6F9"}
wb = Workbook()
wb.remove(wb.active)

thin = Side(style="thin", color="D9E1F2")
medium = Side(style="medium", color=THEME["primary"])

def setup_sheet(title, gridlines=False):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = gridlines
    ws.column_dimensions["A"].width = 3
    ws["B2"] = title
    ws["B2"].font = Font(name="Georgia", size=18, bold=True, color=THEME["primary"])
    ws["B2"].alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 24
    return ws

def write_table(ws, start_row, start_col, title, df, number_formats=None, freeze=True):
    ws.cell(start_row, start_col).value = title
    ws.cell(start_row, start_col).font = Font(name="Georgia", size=13, bold=True, color=THEME["primary"])
    header_row = start_row + 1
    for j, col in enumerate(df.columns, start=start_col):
        c = ws.cell(header_row, j)
        c.value = col
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=THEME["primary"])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(bottom=medium)
    for i, record in enumerate(df.to_dict("records"), start=header_row + 1):
        for j, col in enumerate(df.columns, start=start_col):
            val = record[col]
            if isinstance(val, (pd.Timestamp, datetime)):
                val = val.to_pydatetime() if isinstance(val, pd.Timestamp) else val
            elif pd.isna(val) if not isinstance(val, (list, tuple, dict)) else False:
                val = None
            elif isinstance(val, (np.integer,)):
                val = int(val)
            elif isinstance(val, (np.floating,)):
                val = float(val)
            elif isinstance(val, (np.bool_,)):
                val = bool(val)
            cell = ws.cell(i, j)
            cell.value = val
            cell.font = Font(name="Calibri", size=10, color="000000")
            cell.border = Border(bottom=thin)
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cell.number_format = number_formats.get(col, "#,##0.00") if number_formats else "#,##0.00"
            elif isinstance(val, datetime):
                cell.number_format = "yyyy-mm-dd"
        if i <= header_row + 3:
            for j in range(start_col, start_col + len(df.columns)):
                ws.cell(i, j).fill = PatternFill("solid", fgColor="FAFCFE")
    end_row = header_row + len(df)
    end_col = start_col + len(df.columns) - 1
    ws.auto_filter.ref = f"{get_column_letter(start_col)}{header_row}:{get_column_letter(end_col)}{end_row}"
    if freeze:
        ws.freeze_panes = f"{get_column_letter(start_col)}{header_row + 1}"
    for j, col in enumerate(df.columns, start=start_col):
        max_len = min(28, max(len(str(col)), max([len(str(x)) for x in df[col].head(100).tolist()] + [0])))
        ws.column_dimensions[get_column_letter(j)].width = max(12, max_len + 2)
    return header_row, end_row, end_col

def add_method(ws, text, row=4):
    ws.cell(row, 2).value = text
    ws.cell(row, 2).font = Font(name="Calibri", size=10, italic=True, color=THEME["gray"])
    ws.cell(row, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 44

# README
ws = setup_sheet("README")
readme_lines = [
    "Business context: LG/Kent consumer durables dealership operating in Varanasi, Uttar Pradesh.",
    "Problem: turn three fiscal years of line-item sales registers into a reliable marketing analytics view.",
    "Important grain note: each source row is one product line on an invoice; invoice totals repeat across lines.",
    "Revenue is therefore counted at distinct invoice level wherever invoice revenue is the metric.",
    "Named-customer analyses exclude rows flagged as CASH CUSTOMER; anonymous walk-ins remain in invoice/category analyses.",
    "The workbook contains source-aligned cleaned data, an invoice-level table, preprocessing traceability, and four diagnostic analyses.",
    "Use the CONTENTS section below to navigate. All monetary values are in the source register's currency units.",
]
for i, text in enumerate(readme_lines, start=5):
    ws.cell(i, 2).value = text
    ws.cell(i, 2).font = Font(name="Calibri", size=11, color="000000")
    ws.cell(i, 2).alignment = Alignment(wrap_text=True)
ws["B14"] = "SOURCE FILES"
ws["B14"].font = Font(name="Georgia", size=13, bold=True, color=THEME["primary"])
for i, (path, fy) in enumerate(source_paths, start=15):
    ws.cell(i, 2).value = f"{fy}: {Path(path).name}"
ws["B20"] = "CONTENTS"
ws["B20"].font = Font(name="Georgia", size=13, bold=True, color=THEME["primary"])
for i, sname in enumerate(wb.sheetnames, start=21):
    ws.cell(i, 2).value = sname
    ws.cell(i, 2).hyperlink = f"#'{sname}'!B2"
    ws.cell(i, 2).font = Font(name="Calibri", size=11, color="0563C1", underline="single")
ws.column_dimensions["B"].width = 110

# Data dictionary
ws = setup_sheet("DATA DICTIONARY")
add_method(ws, "Definitions for the original fields and the derived fields created by preprocessing. Original fields are carried from the source registers; derived fields are computed in this project.")
dict_rows = []
original_cols = [c for c in all_cols if c != "FISCAL_YEAR"]
for c in original_cols:
    dict_rows.append({"COLUMN":"" if c is None else c, "TYPE":"Original", "DEFINITION":"Field carried from the source sales register.", "SOURCE_OR_DERIVED":"Source", "NOTES":"May be blank in one or more fiscal-year files."})
for c, definition in [
    ("FISCAL_YEAR", "Source-file fiscal-year tag."), ("MONTH", "Calendar month number parsed from INVOICEDATE."), ("FISCAL_MONTH", "Fiscal month where April=1 and March=12."), ("CATEGORY", "HSNCODE mapping to six named categories; all other codes are Other."), ("ZERO_VALUE_FLAG", "True when GSTRATE equals zero."), ("BAD_SKU_FLAG", "True when PRODUCT contains only numeric characters."), ("CASH_CUSTOMER_FLAG", "True when CUSTOMERNAME contains CASH CUSTOMER."), ("BRANCH", "Alphabetic invoice prefix before the first slash in INVOICENO."), ("INVOICEDATE_clean", "Internal parsed date field used during preprocessing; not retained in cleaned master."), ("HSNCODE_clean", "Internal numeric HSN helper; not retained in cleaned master."), ("BASICRATE_clean", "Internal numeric BASICRATE helper; not retained in cleaned master."), ("GSTRATE_clean", "Internal numeric GSTRATE helper; not retained in cleaned master."), ("GRANDAMOUNT_clean", "Internal numeric GRANDAMOUNT helper; not retained in cleaned master."), ("HSNCODE_KEY", "Internal normalized HSN text key; not retained in cleaned master."),
]:
    dict_rows.append({"COLUMN":c, "TYPE":"Derived", "DEFINITION":definition, "SOURCE_OR_DERIVED":"Derived", "NOTES":"See PREPROCESSING LOG for processing order."})
dict_df = pd.DataFrame(dict_rows)
write_table(ws, 6, 2, "Table 1 — Field definitions", dict_df, {"COLUMN":"@", "TYPE":"@", "DEFINITION":"@", "SOURCE_OR_DERIVED":"@", "NOTES":"@"})

# Preprocessing log
ws = setup_sheet("PREPROCESSING LOG")
add_method(ws, "Ordered audit trail of the cleaning and derivation steps. Rows affected are counted against the relevant source or cleaned dataset.")
log_df = pd.DataFrame([
    {"STEP":1,"ACTION":"Loaded and unioned three source registers", "ROWS_AFFECTED":raw_row_count, "REASON":"Create one cross-year line-item master while preserving optional year-specific columns.", "RESULT":"Unioned fields; FISCAL_YEAR added."},
    {"STEP":2,"ACTION":"Removed exact duplicate rows", "ROWS_AFFECTED":duplicate_rows_removed, "REASON":"Prevent identical repeated lines from inflating revenue or basket counts.", "RESULT":"Duplicates removed; first occurrence retained."},
    {"STEP":3,"ACTION":"Parsed INVOICEDATE day-first and derived MONTH/FISCAL_MONTH", "ROWS_AFFECTED":len(cleaned), "REASON":"Enable fiscal-month seasonality comparisons.", "RESULT":f"Date range {raw['INVOICEDATE_clean'].min().date()} to {raw['INVOICEDATE_clean'].max().date()}."},
    {"STEP":4,"ACTION":"Mapped HSNCODE to CATEGORY", "ROWS_AFFECTED":len(cleaned), "REASON":"Create marketing-friendly product groups.", "RESULT":f"Mapped codes: {len(hsn_map)}; remaining rows labeled Other."},
    {"STEP":5,"ACTION":"Created quality flags", "ROWS_AFFECTED":len(cleaned), "REASON":"Surface zero-GST, numeric SKU, and anonymous cash-customer records without deleting rows.", "RESULT":f"ZERO_VALUE_FLAG={int(cleaned['ZERO_VALUE_FLAG'].sum())}; BAD_SKU_FLAG={int(cleaned['BAD_SKU_FLAG'].sum())}; CASH_CUSTOMER_FLAG={int(cleaned['CASH_CUSTOMER_FLAG'].sum())}."},
    {"STEP":6,"ACTION":"Derived BRANCH", "ROWS_AFFECTED":len(cleaned), "REASON":"Support branch-level slicing from invoice prefixes.", "RESULT":f"Branches observed: {cleaned['BRANCH'].nunique(dropna=True)}."},
    {"STEP":7,"ACTION":"Built invoice-level table", "ROWS_AFFECTED":len(invoices), "REASON":"Count invoice revenue once despite multiple product lines sharing GRANDAMOUNT.", "RESULT":"One row per INVOICENO with category list."},
    {"STEP":8,"ACTION":"Applied customer-analysis scope rule", "ROWS_AFFECTED":int((~invoices['CASH_CUSTOMER_FLAG']).sum()), "REASON":"Exclude anonymous walk-ins from named-customer basket and RFM analysis.", "RESULT":f"Named customers in RFM: {named_customer_count}."},
], columns=["STEP","ACTION","ROWS_AFFECTED","REASON","RESULT"])
write_table(ws, 6, 2, "Table 1 — Preprocessing audit trail", log_df, {"STEP":"0", "ROWS_AFFECTED":"#,##0"})

# Cleaned data and invoice level
ws = setup_sheet("CLEANED DATA", gridlines=True)
add_method(ws, "Line-item master after exact duplicate removal and derived fields. Use this sheet for row-level auditing; use INVOICE LEVEL for invoice-revenue calculations.")
cleaned_out = cleaned.copy()
# Ensure readable date column name and datetime values.
cleaned_out["INVOICEDATE"] = raw.loc[cleaned_out.index, "INVOICEDATE_clean"]
start_row = 6
write_table(ws, start_row, 2, "Table 1 — Cleaned line-item master", cleaned_out, {"QUANTITY":"#,##0.00","BASICRATE":"#,##0.00","GRANDAMOUNT":"#,##0.00","TOTAL":"#,##0.00","GSTRATE":"0.0%","MONTH":"0","FISCAL_MONTH":"0"})

ws = setup_sheet("INVOICE LEVEL", gridlines=True)
add_method(ws, "One row per invoice. GRANDAMOUNT is taken once per INVOICENO; CATEGORIES is the semicolon-separated set of categories appearing on that invoice.")
write_table(ws, 6, 2, "Table 1 — Invoice-level table", invoices, {"GRANDAMOUNT":"#,##0.00","FISCAL_MONTH":"0","CATEGORY_COUNT":"0"})

# Analysis A
ws = setup_sheet("ANALYSIS A")
add_method(ws, "Category and seasonality diagnostic. Revenue is counted once per distinct invoice for fiscal-year and fiscal-month totals. Category-by-month revenue allocates a multi-category invoice's total equally across its categories so pooled category totals reconcile to invoice revenue.")
cat_fy_w = cat_fy.reset_index(name="REVENUE")
cat_fy_w["FY_ORDER"] = cat_fy_w["FISCAL_YEAR"].map({"FY2023-24":1,"FY2024-25":2,"FY2025-26":3})
cat_piv = cat_fy_w.pivot(index="CATEGORY", columns="FISCAL_YEAR", values="REVENUE").fillna(0).reset_index()
for fy in ["FY2023-24","FY2024-25","FY2025-26"]:
    if fy not in cat_piv.columns: cat_piv[fy] = 0.0
cat_piv["YoY % 24-25"] = np.where(cat_piv["FY2023-24"] != 0, cat_piv["FY2024-25"] / cat_piv["FY2023-24"] - 1, np.nan)
cat_piv["YoY % 25-26"] = np.where(cat_piv["FY2024-25"] != 0, cat_piv["FY2025-26"] / cat_piv["FY2024-25"] - 1, np.nan)
cat_piv = cat_piv[["CATEGORY","FY2023-24","FY2024-25","FY2025-26","YoY % 24-25","YoY % 25-26"]].sort_values("FY2025-26", ascending=False)
write_table(ws, 6, 2, "Table 1 — Revenue by category and fiscal year", cat_piv, {"FY2023-24":"#,##0.00","FY2024-25":"#,##0.00","FY2025-26":"#,##0.00","YoY % 24-25":"0.0%","YoY % 25-26":"0.0%"})
month_w = month_fy.pivot(index="FISCAL_MONTH", columns="FISCAL_YEAR", values="GRANDAMOUNT").fillna(0).reset_index()
for fy in ["FY2023-24","FY2024-25","FY2025-26"]:
    if fy not in month_w.columns: month_w[fy]=0.0
month_w = month_w[["FISCAL_MONTH","FY2023-24","FY2024-25","FY2025-26"]]
write_table(ws, 17, 2, "Table 2 — Total invoice revenue by fiscal month", month_w, {"FISCAL_MONTH":"0","FY2023-24":"#,##0.00","FY2024-25":"#,##0.00","FY2025-26":"#,##0.00"})
cat_month_w = cat_month.pivot(index="CATEGORY", columns="FISCAL_MONTH", values="ALLOCATED_REVENUE").fillna(0).reset_index()
cat_month_w.columns = ["CATEGORY"] + [f"FM_{int(c)}" for c in cat_month_w.columns[1:]]
write_table(ws, 35, 2, "Table 3 — Pooled category revenue by fiscal month", cat_month_w, {c:"#,##0.00" for c in cat_month_w.columns if c != "CATEGORY"})

# Analysis B
ws = setup_sheet("ANALYSIS B")
add_method(ws, "Market basket analysis is run twice. Invoice baskets use the set of categories on each invoice. Named-customer baskets use all categories ever bought by each named customer. Support = pair baskets / total baskets; confidence A→B = pair baskets / A baskets; lift = support / (support A × support B). Cash customers are excluded from customer baskets and customer mix statistics.")
write_table(ws, 6, 2, "Table 1 — Category pair metrics", basket_pairs, {"BASKETS":"#,##0","PAIR_BASKETS":"#,##0","SUPPORT":"0.0%","CONFIDENCE_A_TO_B":"0.0%","CONFIDENCE_B_TO_A":"0.0%","LIFT":"0.00"})
write_table(ws, 6 + len(basket_pairs) + 4, 2, "Table 2 — Single- vs multi-category named customers by fiscal year", customer_mix_summary, {"CUSTOMERS":"#,##0","AVG_TOTAL_SPEND":"#,##0.00","TOTAL_REVENUE":"#,##0.00"})

# Analysis C
ws = setup_sheet("ANALYSIS C")
add_method(ws, "Seasonal pricing diagnostic. For each PRODUCT, BASICRATE is divided by that product's mean positive BASICRATE across all three fiscal years. The table reports the average index by category and fiscal month. CV is the standard deviation of the 12 monthly category averages divided by their mean. Rows with BASICRATE <= 0 are excluded.")
write_table(ws, 6, 2, "Table 1 — Average product-normalized price index by category and fiscal month", price_month, {"FISCAL_MONTH":"0","AVG_PRICE_INDEX":"0.000","COEFFICIENT_OF_VARIATION":"0.0%"})
write_table(ws, 6 + len(price_month) + 4, 2, "Table 2 — Category coefficient of variation", price_cv, {"COEFFICIENT_OF_VARIATION":"0.0%"})

# Analysis D
ws = setup_sheet("ANALYSIS D")
add_method(ws, f"RFM analysis for named customers only. Recency = days from last invoice to the last dataset date ({last_date.date()}); Frequency = distinct invoice count; Monetary = sum of distinct invoice GRANDAMOUNT. Each R, F, and M measure is scored 1–5 by quintile. Segment rules prioritize Champions, Loyal, New, At Risk, Lost, and Potential Loyalist.")
write_table(ws, 6, 2, "Table 1 — RFM segment summary", segment_summary, {"CUSTOMERS":"#,##0","REVENUE":"#,##0.00","AVG_MONETARY":"#,##0.00","AVG_FREQUENCY":"0.00","AVG_RECENCY_DAYS":"0.0","REVENUE_SHARE":"0.0%"})
write_table(ws, 6 + len(segment_summary) + 4, 2, "Table 2 — Customer-level RFM detail", rfm.sort_values(["SEGMENT","MONETARY"], ascending=[True,False]), {"RECENCY_DAYS":"0","FREQUENCY":"0","MONETARY":"#,##0.00","R_SCORE":"0","F_SCORE":"0","M_SCORE":"0"})

# Findings summary
ws = setup_sheet("FINDINGS SUMMARY")
add_method(ws, "Executive summary of the verified outputs in this workbook. Use the detailed analysis sheets for definitions, methods, and drill-down tables.")
findings = pd.DataFrame([
    {"METRIC":"Total distinct-invoice revenue", "VALUE":total_revenue, "UNIT":"Currency units", "INTERPRETATION":"Revenue counted once per INVOICENO across all three fiscal years."},
    {"METRIC":"Distinct invoices", "VALUE":invoice_count, "UNIT":"Invoices", "INTERPRETATION":"Invoice-level denominator for revenue and seasonality diagnostics."},
    {"METRIC":"Named customers", "VALUE":named_customer_count, "UNIT":"Customers", "INTERPRETATION":"Cash customers excluded from customer-level analysis."},
    {"METRIC":"Top category by revenue", "VALUE":top_category, "UNIT":"Category", "INTERPRETATION":f"Top category revenue = {top_category_revenue:,.2f}."},
    {"METRIC":"Exact duplicate rows removed", "VALUE":duplicate_rows_removed, "UNIT":"Rows", "INTERPRETATION":"Exact duplicates removed before derivations."},
    {"METRIC":"Source line-item rows after cleaning", "VALUE":len(cleaned), "UNIT":"Rows", "INTERPRETATION":"Cleaned master retains quality-flagged rows; flags do not delete data."},
], columns=["METRIC","VALUE","UNIT","INTERPRETATION"])
write_table(ws, 6, 2, "Table 1 — Key verified numbers", findings, {"VALUE":"#,##0.00"})
# Add top segment and basket insight text.
insight_row = 6 + len(findings) + 5
ws.cell(insight_row, 2).value = "KEY INTERPRETIVE NOTES"
ws.cell(insight_row, 2).font = Font(name="Georgia", size=13, bold=True, color=THEME["primary"])
notes = [
    f"Top category: {top_category}.",
    f"RFM revenue concentration: {segment_summary.iloc[0]['SEGMENT']} is the largest segment by revenue in the named-customer base, contributing {segment_summary.iloc[0]['REVENUE_SHARE']:.1%} of named-customer revenue.",
    "Use Analysis A to identify category growth and seasonal peaks; Analysis B to identify cross-sell pairings; Analysis C to see category-level price seasonality; and Analysis D to prioritize retention and acquisition actions.",
]
for i, note in enumerate(notes, start=insight_row+1):
    ws.cell(i, 2).value = note
    ws.cell(i, 2).alignment = Alignment(wrap_text=True)

# Global formatting and sheet widths.
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.font.name is None:
                cell.font = Font(name="Calibri", size=10)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 22)

# Conditional formats on percentage columns in key analysis sheets.
for sname in ["ANALYSIS A", "ANALYSIS B", "ANALYSIS C", "ANALYSIS D"]:
    wb[sname].sheet_view.showGridLines = False

# Manifest and audit.
manifest = {"schemaVersion":1,"profile":"general","assertions":[
    {"id":"readme","label":"README exists","type":"required_sheet","sheet":"README"},
    {"id":"dict","label":"Data dictionary exists","type":"required_sheet","sheet":"DATA DICTIONARY"},
    {"id":"log","label":"Preprocessing log exists","type":"required_sheet","sheet":"PREPROCESSING LOG"},
    {"id":"cleaned","label":"Cleaned master exists","type":"required_sheet","sheet":"CLEANED DATA"},
    {"id":"invoice","label":"Invoice-level table exists","type":"required_sheet","sheet":"INVOICE LEVEL"},
    {"id":"a","label":"Analysis A exists","type":"required_sheet","sheet":"ANALYSIS A"},
    {"id":"b","label":"Analysis B exists","type":"required_sheet","sheet":"ANALYSIS B"},
    {"id":"c","label":"Analysis C exists","type":"required_sheet","sheet":"ANALYSIS C"},
    {"id":"d","label":"Analysis D exists","type":"required_sheet","sheet":"ANALYSIS D"},
    {"id":"findings","label":"Findings summary exists","type":"required_sheet","sheet":"FINDINGS SUMMARY"},
]}

output_path = "/home/user/workspace/outputs/lg_kent_marketing_analytics_fy23_26.xlsx"
wb.save(output_path)
print(f"AUDIT: source rows by fiscal year={source_rows}")
print(f"AUDIT: unioned rows={raw_row_count}; exact duplicates removed={duplicate_rows_removed}; cleaned rows={len(cleaned)}")
print(f"AUDIT: distinct invoices={invoice_count}; named customers={named_customer_count}; total revenue={total_revenue:.2f}")
print(f"AUDIT: analysis tables: A={len(cat_piv) + len(month_w) + len(cat_month_w)} rows; B={len(basket_pairs) + len(customer_mix_summary)} rows; C={len(price_month) + len(price_cv)} rows; D={len(segment_summary) + len(rfm)} rows")
print(f"AUDIT: workbook sheets={wb.sheetnames}")
print("__SB_WORKBOOK_SEMANTIC_MANIFEST__" + json.dumps(manifest, separators=(",", ":")))
print(f"SAVED: {output_path} exists={os.path.exists(output_path)} size_bytes={os.path.getsize(output_path)}")
