# LLM Usage Log

Disclosure of AI assistance in this project, per the submission requirement
that tools used and logs of LLM usage be included.

---

## Tools

| LLM tool | Role |
|---|---|
| **Better Analyst** (app.betteranalyst.com) | Agentic data analysis. Wrote and executed Python to produce workbook v1 and corrected v2. Fully documented in `better_analyst/`. |
| **Claude** (Anthropic) | Analytical assistant. Scoping, verification, error detection, secondary research, drafting. |

---

## What Claude was used for

**Feasibility screening.** Before any modelling, all sixteen techniques from
the course were screened against the raw data. Four empirical conditions were
tested — basket structure, price dispersion, repeat-purchase rate and
promotional-field completeness — and each technique classed as fully supported,
partially supported, or not feasible. Six techniques (conjoint, market-share
simulation, advertising effectiveness and planning, promotion and display
planning, network/viral analysis) were excluded on evidence rather than
performed badly. This screening is reported in the report's Section 4.

**Method selection.** Choosing the four chained modules and the rationale for
each: line-level rather than invoice-level revenue; a like-for-like comparison
window; lift at two basket levels; a product-normalised price index; and the
explicit rejection of the standard FMCG reading of RFM.

**Independent verification.** Every figure Better Analyst produced was
recomputed from the raw files in a separate Python environment. This is what
caught the category double-counting error.

**Error detection.** Two errors were found, each of which would have inverted
the study's conclusions:

1. *Category revenue double-counting* (Better Analyst's Analysis A) — caught by
   a reconciliation test showing the parts exceeded the whole.
2. *FY2024-25 data completeness gap* — caught by testing monthly invoice counts
   for plausibility. February holds 298 invoices and March 35, against ~800 and
   ~1,100 in adjacent years. This had manufactured the apparent "collapse and
   recovery" narrative; on a like-for-like basis the business simply eroded
   7.9%.

**Secondary research.** Replacement-cycle and warranty evidence for Indian
durables (refrigerator and washing-machine warranty terms, television
replacement age, the 2025 iFOREST refrigerant-refilling finding), Varanasi
climate records, and post-Kashi-Vishwanath-Corridor tourism growth. Sources
were checked; where figures were ambiguous — Varanasi's population varies from
~1.2m to ~3.9m depending on boundary definition — a range is cited rather than
false precision.

**Drafting.** The report, the analysis and workbook-building scripts, and the
presentation script.

---

## Verification standard applied

**No quantitative claim in the report, workbook or deck rests on an LLM
assertion.** Every figure is computed by code in `code/` and can be reproduced
by running `python analysis.py`. Where the two independent computations
disagreed, the raw data was treated as authoritative and the discrepancy
investigated — which is how both errors above were found.

Two runtime assertions are built into `code/analysis.py` so the checks are
enforced rather than merely claimed:

```python
assert (varies <= 1).all(), "GRANDAMOUNT varies within an invoice"
assert 0.99 <= ratio <= 1.01, f"{fy} fails reconciliation at {ratio:.4f}"
```

---

## Honest note on limitations

One claim that appeared in an earlier draft — a lift value asserting the two
washing-machine variants are substitutes — could not be reproduced from any
verified output and was **removed** rather than retained. It is recorded here
because the standard applied throughout was that an unverifiable figure does
not go in the report, regardless of how plausible it sounds.
