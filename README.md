# Dealership Sales Analytics: Category Decline & the Cross-Sell Gap

**Data Driven Marketing · IIM Nagpur · Study Group 02, Section A**

A three-year marketing-analytics study of a multi-branch consumer-durables
dealership in Varanasi (LG white goods and Kent water purifiers), built on
67,334 real invoice lines. Two data errors, caught by reconciliation, would each
have inverted the conclusions.

![Python](https://img.shields.io/badge/Python-pandas%20%7C%20numpy%20%7C%20openpyxl-3776AB)
![Methods](https://img.shields.io/badge/methods-market%20basket%20%7C%20RFM%20%7C%20price%20index-555)
![Excel](https://img.shields.io/badge/Excel-15--sheet%20workbook-217346)

---

## The problem

The brief was to find a business problem actually observable in the firm's own
data, not to apply a taught technique for its own sake. The reported numbers
showed a sharp fall and recovery. The real story turned out to be different.

## Two corrections that changed the answer

**1. Line-level revenue, not invoice totals.** Category is a line attribute;
`GRANDAMOUNT` is an invoice total. Attributing invoice totals to each category
double-counted every multi-category invoice (40–43% of invoices): four categories
alone summed to ₹472.0m against a true FY2023-24 total of ₹435.5m. Rebuilt as
`LINE_REVENUE = TOTAL + CGST + SGST + IGST`, which reconciles to 100.1% of invoice
totals in all three years (asserted in code). Air Conditioner moved from
"third-largest and declining" to *largest and the only growing category*.

**2. Like-for-like window, not full years.** FY2024-25 February and March hold 298
and 35 invoices against ~800 and ~1,100 in adjacent years: a truncated export, not
a trading collapse. On a comparable April–January window the "−13.7% then +9.3%"
V-shape disappears into a **steady 7.9% erosion over two years**.

## What the data says

| Finding | Evidence |
|---|---|
| Only one category is growing | Air Conditioner +16.8% vs FY2023-24; the other six fell 10.2–32.0%, consistent with long replacement cycles |
| The growth engine is the least cross-sold | AC lift **below 1.0** against every other category (0.475–0.849 per invoice, 0.707–0.960 per customer); TV, refrigerator and washing machines cluster at 1.28–1.61 |
| Cross-sell is where the value is | Multi-category customers are worth **3.4–3.7×** single-category ones, yet single-category share rose from 64.1% to 69.4% |

That is the mechanism by which AC revenue can grow 16.8% while like-for-like revenue
falls 7.9%. The workbook turns it into 9 findings, 7 prioritised recommendations and
10 stated limitations.

## Methods

Data cleaning (12 logged steps) · reconciliation checks · category and seasonality
analysis · market-basket analysis (support, confidence, lift) at invoice and customer
level · product-normalised price index · RFM segmentation, recalibrated for durables ·
first-pass exploration in an agentic AI analytics tool, whose methodological error was
found, documented and corrected

## Repository contents

```
code/
  analysis.py             Reproduces the full analysis and prints every figure in the report
  build_workbook.py       Builds the full workbook locally (live formulas, zero errors)
  requirements.txt
  RUN_LOG.txt             Console output of a complete verified run
  TOOL_LOG.md             Every tool used, what it produced, and the two errors caught
  llm_usage_log.md        AI-assistance disclosure and the verification standard applied
  better_analyst/         The AI analytics tool's own output, kept as an audit trail:
    generated_code_v1.txt         first-pass code, with the category-revenue error left in and annotated
    correction_category_revenue.py  the fix: line-level revenue that reconciles to invoice totals
    final_workbook_build_v3.py    the tool's final workbook build (like-for-like, gap check, findings)
    session_notes.md              brief, outputs, the error, the correction
analysis/
  LG_Kent_Marketing_Analytics_FY2324_FY2526.xlsx   15 analysis sheets: data dictionary,
                                                   cleaning log, gap check, analyses A–D,
                                                   findings, recommendations, limitations
report/
  DDM_Final_Report.pdf
```

## Run it

```bash
cd code
pip install -r requirements.txt
python analysis.py        # ~2–3 min; also regenerates the workbook
```

## Data and privacy

The dealership's sales registers and all row-level transaction data are **not
published**. The workbook here contains the analysis sheets only. The scripts expect
the registers in `code/` or `data/`. Running them locally rebuilds the full workbook,
including the row-level sheets with buyer names, so that output stays private (it is
listed in `.gitignore`).

**Team:** Aishwarya Pratap Singh, Adarsh, Aniket Mohankar, Anshil Seth, Deepshika Sidar
