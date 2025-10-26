# AMEISE Project Plan

This repository contains the AMEISE-based project schedule, cost estimation, and an automated Excel workbook builder.

Contents
- data/AMEISE-Schedule-v2-grid.csv — weekly Gantt-like grid
- data/Cost-Estimation-v2.csv — cost plan, budget cap, and slack
- docs/AMEISE-Estimation-OptionA.md — estimation details (26.36 PM, 8.67 months, AvgDev≈3.04)
- docs/AMEISE-Schedule-v2.md — narrative schedule and AMEISE gates
- scripts/build_excel.py — builds AMEISE-Plan.xlsx using pandas + openpyxl
- .github/workflows/build-excel.yml — GitHub Action to generate and commit Excel, and upload it as an artifact

How to use
1) Edit the CSVs in data/.
2) Push to main. The workflow will generate AMEISE-Plan.xlsx at repo root and attach it as an artifact.
3) Excel includes:
   - Sheet "Schedule" — color-coded weekly plan
   - Sheet "Cost Estimation" — totals, budget cap check, slack, conditional formatting
   - Sheet "OptionA" — COCOMO-like metrics and effort distribution

Budget
- Current cap: 225,000 EUR (can be changed via repository variable BUDGET_CAP or by editing scripts/build_excel.py).