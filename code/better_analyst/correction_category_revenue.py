"""
Fix for the category-revenue reconciliation error found in v1 (see TOOL_LOG.md).

Source: generated and run in Better Analyst (app.betteranalyst.com), session
"Marketing Analytics Workbook Preview", 2026-09-04. Exported verbatim; only this header was added.
Input paths point to Better Analyst's sandbox: /home/user/workspace/1788497159764_SALE_DATA_2425_NEW.xlsx, /home/user/workspace/1788497160021_SALE_DATA.xlsx, /home/user/workspace/1788497160280_SALE_DATA_2324.xlsx, /home/user/workspace/outputs/lg_kent_marketing_analytics_fy23_26_line_revenue_corrected_v2.xlsx, /home/user/workspace/outputs/runs/4ddcd727-9d46-4e45-91e7-62a42effe440/attempts/b32ee1dd-6c2d-497a-a8f1-78543fbb27ee/lg_kent_marketing_analytics_fy23_26.xlsx.
Change them to your local data folder before running.
"""

from pathlib import Path
from itertools import combinations
from copy import copy
import os, re, json
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

source_book = "/home/user/workspace/outputs/runs/4ddcd727-9d46-4e45-91e7-62a42effe440/attempts/b32ee1dd-6c2d-497a-a8f1-78543fbb27ee/lg_kent_marketing_analytics_fy23_26.xlsx"
if not Path(source_book).is_file():
    raise ValueError(f"Missing generated workbook: {source_book}")
print(f"AUDIT: source workbook={source_book}")

input_paths = [
    ("/home/user/workspace/1788497160280_SALE_DATA_2324.xlsx", "FY2023-24"),
    ("/home/user/workspace/1788497159764_SALE_DATA_2425_NEW.xlsx", "FY2024-25"),
    ("/home/user/workspace/1788497160021_SALE_DATA.xlsx", "FY2025-26"),
]
for pth, _fy in input_paths:
    if not Path(pth).is_file():
        raise ValueError(f"Missing input workbook: {pth}")

# Reconstruct the cleaned line-level source deterministically from the three original registers.
frames = []
for pth, fy in input_paths:
    df = pd.read_excel(pth, sheet_name="Sheet1")
    df.columns = [str(c).strip() for c in df.columns]
    df["FISCAL_YEAR"] = fy
    frames.append(df)
all_cols = []
for df in frames:
    for c in df.columns:
        if c not in all_cols:
            all_cols.append(c)
raw = pd.concat([df.reindex(columns=all_cols) for df in frames], ignore_index=True)
pre_dedupe = len(raw)
raw = raw.drop_duplicates(keep="first").reset_index(drop=True)
removed = pre_dedupe - len(raw)
raw["INVOICEDATE_clean"] = pd.to_datetime(raw["INVOICEDATE"], dayfirst=True, errors="coerce")
raw["HSNCODE_clean"] = pd.to_numeric(raw["HSNCODE"], errors="coerce")
for c in ["TOTAL", "SGSTAMT", "CGSTAMT", "IGSTAMT", "GRANDAMOUNT", "BASICRATE"]:
    raw[c + "_clean"] = pd.to_numeric(raw[c], errors="coerce").fillna(0.0)
raw["MONTH"] = raw["INVOICEDATE_clean"].dt.month
raw["FISCAL_MONTH"] = ((raw["MONTH"] - 4) % 12) + 1
hsn_map = {841510:"Air Conditioner", 852872:"Television", 841810:"Refrigerator", 84501100:"Washing Machine Automatic", 84501200:"Washing Machine Semi-Automatic", 842121:"Water Purifier RO"}
raw["CATEGORY"] = raw["HSNCODE_clean"].map(hsn_map).fillna("Other")
raw["LINE_REVENUE"] = raw["TOTAL_clean"] + raw["SGSTAMT_clean"] + raw["CGSTAMT_clean"] + raw["IGSTAMT_clean"]
raw["BRANCH"] = raw["INVOICENO"].astype("string").fillna("").str.split("/", n=1).str[0].str.strip()

# Corrected Analysis A calculations.
cat_year = raw.groupby(["CATEGORY", "FISCAL_YEAR"], dropna=False)["LINE_REVENUE"].sum().unstack(fill_value=0)
for fy in ["FY2023-24", "FY2024-25", "FY2025-26"]:
    if fy not in cat_year.columns:
        cat_year[fy] = 0.0
cat_year = cat_year[["FY2023-24", "FY2024-25", "FY2025-26"]]
cat_year["YoY % 24-25"] = np.where(cat_year["FY2023-24"] != 0, cat_year["FY2024-25"] / cat_year["FY2023-24"] - 1, np.nan)
cat_year["YoY % 25-26"] = np.where(cat_year["FY2024-25"] != 0, cat_year["FY2025-26"] / cat_year["FY2024-25"] - 1, np.nan)
cat_year = cat_year.sort_values("FY2023-24", ascending=False)

# Distinct-invoice totals are retained only for fiscal-month seasonality.
inv = raw.sort_values(["INVOICENO", "INVOICEDATE_clean"]).drop_duplicates("INVOICENO", keep="first")
monthly_inv = inv.groupby(["FISCAL_MONTH", "FISCAL_YEAR"], dropna=False)["GRANDAMOUNT_clean"].sum().unstack(fill_value=0)
for fy in ["FY2023-24", "FY2024-25", "FY2025-26"]:
    if fy not in monthly_inv.columns:
        monthly_inv[fy] = 0.0
monthly_inv = monthly_inv[["FY2023-24", "FY2024-25", "FY2025-26"]].reindex(range(1,13), fill_value=0)
pooled = raw.groupby(["CATEGORY", "FISCAL_MONTH"], dropna=False)["LINE_REVENUE"].sum().unstack(fill_value=0).reindex(columns=range(1,13), fill_value=0)
pooled = pooled.loc[cat_year.index]

cat_recon = cat_year[["FY2023-24", "FY2024-25", "FY2025-26"]].sum(axis=0)
inv_totals = inv.groupby("FISCAL_YEAR")["GRANDAMOUNT_clean"].sum().reindex(["FY2023-24", "FY2024-25", "FY2025-26"], fill_value=0)
recon_pct = cat_recon / inv_totals
print(f"AUDIT: raw_rows={pre_dedupe} duplicate_rows_removed={removed} cleaned_rows={len(raw)}")
print(f"AUDIT: line_revenue_totals={cat_recon.to_dict()}")
print(f"AUDIT: distinct_invoice_totals={inv_totals.to_dict()}")
print(f"AUDIT: reconciliation_pct={recon_pct.to_dict()}")
expected = {"FY2023-24":139511416, "FY2024-25":149978183, "FY2025-26":166097410}
for fy, target in expected.items():
    observed = float(cat_year.loc["Air Conditioner", fy])
    if abs(observed - target) > 5000:
        raise ValueError(f"Air Conditioner verification failed for {fy}: observed {observed}, expected about {target}")

# Open the generated workbook and preserve all non-A sheets, replacing only the stale Analysis A and updating the cleaned data/log/summary.
wb = openpyxl.load_workbook(source_book, data_only=False)
print(f"AUDIT: inherited sheets={wb.sheetnames}")

# Add LINE_REVENUE to CLEANED DATA using observed header/data positions.
cd = wb["CLEANED DATA"]
header_row = None
for r in range(1, min(cd.max_row, 15) + 1):
    vals = [cd.cell(r, c).value for c in range(1, min(cd.max_column, 80) + 1)]
    if "INVOICENO" in vals and "TOTAL" in vals:
        header_row = r
        break
if header_row is None:
    raise ValueError("Could not locate CLEANED DATA header row")
header_map = {cd.cell(header_row, c).value: c for c in range(1, cd.max_column + 1) if cd.cell(header_row, c).value not in (None, "")}
if "LINE_REVENUE" in header_map:
    line_col = header_map["LINE_REVENUE"]
else:
    line_col = max(header_map.values()) + 1
    cd.cell(header_row, line_col).value = "LINE_REVENUE"
    cd.cell(header_row, line_col).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cd.cell(header_row, line_col).fill = PatternFill("solid", fgColor="1F4E79")
    cd.cell(header_row, line_col).alignment = Alignment(horizontal="center")
first_data = header_row + 1
last_data = first_data + len(raw) - 1
if cd.max_row < last_data:
    raise ValueError(f"CLEANED DATA rows {cd.max_row} shorter than expected {last_data}")
for i, value in enumerate(raw["LINE_REVENUE"].tolist(), start=first_data):
    cd.cell(i, line_col).value = float(value)
    cd.cell(i, line_col).number_format = "#,##0.00"
cd.column_dimensions[get_column_letter(line_col)].width = 16
print(f"AUDIT: CLEANED DATA LINE_REVENUE column={get_column_letter(line_col)} rows={len(raw)}")

# Replace Analysis A with a clean, fully documented sheet.
old_index = wb.sheetnames.index("ANALYSIS A")
del wb["ANALYSIS A"]
ws = wb.create_sheet("ANALYSIS A", old_index)
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws["B2"] = "ANALYSIS A — CATEGORY AND SEASONALITY DIAGNOSTIC"
ws["B2"].font = Font(name="Georgia", size=18, bold=True, color="1F4E79")
ws["B4"] = "Method: Category revenue is calculated from line-level LINE_REVENUE = TOTAL + SGSTAMT + CGSTAMT + IGSTAMT, with blanks treated as zero. GRANDAMOUNT is an invoice-level total and is never allocated to every category on a multi-category invoice; using it for category splits double-counts revenue. Distinct-invoice GRANDAMOUNT is used only for the fiscal-month total-revenue table."
ws["B4"].alignment = Alignment(wrap_text=True, vertical="top")
ws["B4"].font = Font(name="Calibri", size=10, italic=True, color="666666")
ws.row_dimensions[4].height = 56

primary = "1F4E79"
light = "D6E3F0"
header_fill = PatternFill("solid", fgColor=primary)
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
thin = Side(style="thin", color="D1D1D1")

def style_table(sheet, r1, r2, c1, c2):
    for c in range(c1, c2 + 1):
        cell = sheet.cell(r1, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            sheet.cell(r, c).border = Border(left=thin if c == c1 else Side(style=None), right=thin if c == c2 else Side(style=None), top=thin if r == r1 else Side(style=None), bottom=thin)

# Table 1.
r = 7
ws.cell(r, 2).value = "Table 1 — Line-level category revenue by fiscal year"
ws.cell(r, 2).font = Font(name="Georgia", size=13, bold=True, color=primary)
r += 1
headers = ["CATEGORY", "FY2023-24", "FY2024-25", "FY2025-26", "YoY % 24-25", "YoY % 25-26"]
for j, h in enumerate(headers, start=2): ws.cell(r, j).value = h
style_table(ws, r, r + len(cat_year), 2, 7)
for i, (cat, rowv) in enumerate(cat_year.iterrows(), start=r + 1):
    ws.cell(i, 2).value = cat
    for j, fy in enumerate(["FY2023-24", "FY2024-25", "FY2025-26"], start=3):
        ws.cell(i, j).value = float(rowv[fy]); ws.cell(i, j).number_format = "#,##0.00"
    ws.cell(i, 6).value = None if pd.isna(rowv["YoY % 24-25"]) else float(rowv["YoY % 24-25"]); ws.cell(i, 6).number_format = "0.0%"
    ws.cell(i, 7).value = None if pd.isna(rowv["YoY % 25-26"]) else float(rowv["YoY % 25-26"]); ws.cell(i, 7).number_format = "0.0%"
recon_row = r + 1 + len(cat_year)
ws.cell(recon_row, 2).value = "RECONCILIATION — sum of category LINE_REVENUE"
for j, fy in enumerate(["FY2023-24", "FY2024-25", "FY2025-26"], start=3):
    ws.cell(recon_row, j).value = float(cat_recon[fy]); ws.cell(recon_row, j).number_format = "#,##0.00"
ws.cell(recon_row, 6).value = "vs distinct invoice total"
ws.cell(recon_row, 7).value = ""
for j, fy in enumerate(["FY2023-24", "FY2024-25", "FY2025-26"], start=8):
    ws.cell(recon_row, j).value = float(recon_pct[fy]); ws.cell(recon_row, j).number_format = "0.0%"
ws.cell(recon_row + 1, 2).value = "Distinct-invoice GRANDAMOUNT total used for reconciliation only"
for j, fy in enumerate(["FY2023-24", "FY2024-25", "FY2025-26"], start=3):
    ws.cell(recon_row + 1, j).value = float(inv_totals[fy]); ws.cell(recon_row + 1, j).number_format = "#,##0.00"

# Table 2.
r2 = recon_row + 4
ws.cell(r2, 2).value = "Table 2 — Total revenue by fiscal month (distinct invoice GRANDAMOUNT)"
ws.cell(r2, 2).font = Font(name="Georgia", size=13, bold=True, color=primary)
r2 += 1
for j, h in enumerate(["FISCAL_MONTH", "FY2023-24", "FY2024-25", "FY2025-26"], start=2): ws.cell(r2, j).value = h
style_table(ws, r2, r2 + 12, 2, 5)
for i, fm in enumerate(range(1, 13), start=r2 + 1):
    ws.cell(i, 2).value = fm
    for j, fy in enumerate(["FY2023-24", "FY2024-25", "FY2025-26"], start=3):
        ws.cell(i, j).value = float(monthly_inv.loc[fm, fy]); ws.cell(i, j).number_format = "#,##0.00"

# Table 3.
r3 = r2 + 15
ws.cell(r3, 2).value = "Table 3 — Pooled category revenue by fiscal month (line-level LINE_REVENUE)"
ws.cell(r3, 2).font = Font(name="Georgia", size=13, bold=True, color=primary)
r3 += 1
pooled_headers = ["CATEGORY"] + [f"FM_{i}" for i in range(1, 13)]
for j, h in enumerate(pooled_headers, start=2): ws.cell(r3, j).value = h
style_table(ws, r3, r3 + len(pooled), 2, 14)
for i, (cat, rowv) in enumerate(pooled.iterrows(), start=r3 + 1):
    ws.cell(i, 2).value = cat
    for j, fm in enumerate(range(1, 13), start=3):
        ws.cell(i, j).value = float(rowv[fm]); ws.cell(i, j).number_format = "#,##0.00"

for col in range(2, 15): ws.column_dimensions[get_column_letter(col)].width = 18 if col > 2 else 32
ws.freeze_panes = "C9"
print(f"AUDIT: Analysis A tables rows={r3 + len(pooled)} category_rows={len(cat_year)}")

# Add methodology note to preprocessing log.
pl = wb["PREPROCESSING LOG"]
next_row = pl.max_row + 2
while any(isinstance(pl.cell(next_row, c), MergedCell) for c in range(2, 6)):
    next_row += 1
pl.cell(next_row, 2).value = "CORRECTION NOTE"
pl.cell(next_row, 3).value = "Category revenue methodology corrected"
pl.cell(next_row, 4).value = "LINE_REVENUE = TOTAL + SGSTAMT + CGSTAMT + IGSTAMT; blanks treated as zero"
pl.cell(next_row, 5).value = "GRANDAMOUNT is invoice-level and cannot be allocated to each category on multi-category invoices because that double-counts revenue."
for c in range(2, 6):
    pl.cell(next_row, c).alignment = Alignment(wrap_text=True, vertical="top")
    pl.cell(next_row, c).fill = PatternFill("solid", fgColor="FFF3E0")
pl.row_dimensions[next_row].height = 42

# Update stale findings summary statements that named Other as top category or cited old category logic.
fs = wb["FINDINGS SUMMARY"]
for row in fs.iter_rows():
    for cell in row:
        if isinstance(cell.value, str):
            if cell.value.strip() == "Other":
                cell.value = "Air Conditioner"
            elif "Top category: Other." in cell.value:
                cell.value = cell.value.replace("Top category: Other.", "Top category: Air Conditioner.")
            elif "Top category by revenue" in cell.value and "Other" in cell.value:
                cell.value = cell.value.replace("Other", "Air Conditioner")
# Append a traceable correction note.
fs_next_row = fs.max_row + 2
while any(isinstance(fs.cell(fs_next_row, c), MergedCell) for c in range(2, 4)):
    fs_next_row += 1
fs.cell(fs_next_row, 2).value = "CORRECTION NOTE"
fs.cell(fs_next_row, 3).value = "Analysis A category revenue now uses line-level LINE_REVENUE; distinct-invoice GRANDAMOUNT is used only for fiscal-month total revenue."
fs.cell(fs_next_row, 2).font = Font(bold=True, color="C62828")
fs.cell(fs_next_row, 3).alignment = Alignment(wrap_text=True)

# Add workbook-visible LINE_REVENUE dictionary entry if not already present.
dd = wb["DATA DICTIONARY"]
found_line = False
for row in dd.iter_rows():
    for cell in row:
        if cell.value == "LINE_REVENUE": found_line = True
if not found_line:
    rr = dd.max_row + 2
    while any(isinstance(dd.cell(rr, c), MergedCell) for c in range(2, 7)):
        rr += 1
    dd.cell(rr, 2).value = "LINE_REVENUE"
    dd.cell(rr, 3).value = "Derived"
    dd.cell(rr, 4).value = "TOTAL + SGSTAMT + CGSTAMT + IGSTAMT, with blanks treated as zero; used for category-level revenue because category is a line-level attribute."
    dd.cell(rr, 5).value = "Currency units"
    dd.cell(rr, 6).value = "Derived line-level revenue"
    for c in range(2, 7): dd.cell(rr, c).alignment = Alignment(wrap_text=True, vertical="top")

out_path = "/home/user/workspace/outputs/lg_kent_marketing_analytics_fy23_26_line_revenue_corrected_v2.xlsx"
if Path(out_path).resolve() == Path(source_book).resolve():
    raise ValueError("Output would overwrite source workbook")
wb.save(out_path)
size = os.path.getsize(out_path)
print(f"AUDIT: output_sheets={wb.sheetnames}")
print(f"AUDIT: corrected_Air_Conditioner={cat_year.loc['Air Conditioner', ['FY2023-24','FY2024-25','FY2025-26']].to_dict()}")
print(f"SAVED: {out_path} exists=True size_bytes={size}")

# Print complete corrected Analysis A tables for the chat.
print("\n===== CORRECTED ANALYSIS A — CATEGORY REVENUE BY FISCAL YEAR =====")
print("CATEGORY\tFY2023-24\tFY2024-25\tFY2025-26\tYoY % 24-25\tYoY % 25-26")
for cat, rowv in cat_year.iterrows():
    print(f"{cat}\t{rowv['FY2023-24']:.12f}\t{rowv['FY2024-25']:.12f}\t{rowv['FY2025-26']:.12f}\t{rowv['YoY % 24-25']:.12f}\t{rowv['YoY % 25-26']:.12f}")
print(f"RECONCILIATION — sum of category LINE_REVENUE\t{cat_recon['FY2023-24']:.12f}\t{cat_recon['FY2024-25']:.12f}\t{cat_recon['FY2025-26']:.12f}")
print(f"RECONCILIATION % vs distinct-invoice GRANDAMOUNT\t{recon_pct['FY2023-24']:.12f}\t{recon_pct['FY2024-25']:.12f}\t{recon_pct['FY2025-26']:.12f}")
print(f"DISTINCT-INVOICE GRANDAMOUNT TOTAL\t{inv_totals['FY2023-24']:.12f}\t{inv_totals['FY2024-25']:.12f}\t{inv_totals['FY2025-26']:.12f}")
print("\n===== CORRECTED ANALYSIS A — TOTAL REVENUE BY FISCAL MONTH =====")
print("FISCAL_MONTH\tFY2023-24\tFY2024-25\tFY2025-26")
for fm in range(1,13):
    print(f"{fm}\t{monthly_inv.loc[fm,'FY2023-24']:.12f}\t{monthly_inv.loc[fm,'FY2024-25']:.12f}\t{monthly_inv.loc[fm,'FY2025-26']:.12f}")
print("\n===== CORRECTED ANALYSIS A — POOLED CATEGORY REVENUE BY FISCAL MONTH =====")
print("CATEGORY\t" + "\t".join([f"FM_{i}" for i in range(1,13)]))
for cat, rowv in pooled.iterrows():
    print(cat + "\t" + "\t".join([f"{rowv[i]:.12f}" for i in range(1,13)]))

manifest = {"schemaVersion":1,"profile":"general","assertions":[
    {"id":"analysis_a","label":"Corrected Analysis A exists","type":"required_sheet","sheet":"ANALYSIS A"},
    {"id":"cleaned_line_revenue","label":"LINE_REVENUE column exists","type":"required_metric","value":{"sheet":"CLEANED DATA","range":f"{get_column_letter(line_col)}{header_row}"},"allowText":True},
    {"id":"air_conditioner_fy2324","label":"Air Conditioner FY2023-24 corrected revenue","type":"required_metric","value":{"sheet":"ANALYSIS A","range":"D9"},"allowText":False},
    {"id":"reconciliation_note","label":"Reconciliation row exists","type":"required_metric","value":{"sheet":"ANALYSIS A","range":f"B{recon_row}"},"allowText":True}
]}
print("__SB_WORKBOOK_SEMANTIC_MANIFEST__" + json.dumps(manifest, separators=(",", ":")))
