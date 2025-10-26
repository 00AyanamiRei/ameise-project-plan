# AMEISE Estimation — Option A (recalculated and validated)

Inputs
- AFP: 200
- LOC: 9800 → KLOC = 9.8
- Month-to-days: 1 month = 30.4 days

Formulas
- Person-Months (PM) = 2.4 × (KLOC)^1.05 → 26.36 PM
- Duration (months) = 2.5 × (PM)^0.38 → 8.67 months (≈ 263.6 days)
- Avg Developers = PM / Duration → 3.04

Effort distribution (includes Mgmt 7% and Reviews 3%; core=90%)
- Spec 10% (OK), Design 27% (OK), Coding 26% (OK), Testing 18% (OK), Manuals 9% (OK)
- Reviews 3%, Management 7%

Checks
- Under budget: 202,521.5 ≤ 225,000 (slack ≈ 22,478.5)
- Waterfall gates with ≥50% document/requirement thresholds before new phases and integration.
