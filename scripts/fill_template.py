import os
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

# Inputs/env
ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = ROOT / os.getenv("TEMPLATE_PATH", "template/AMEISE-planning-template-v1.95-ip.xlsx")
OUT_PATH = ROOT / "template" / "AMEISE-planning-template-v1.95-ip_filled.xlsx"
BUDGET_CAP = float(os.getenv("BUDGET_CAP", "225000"))

# Option A metrics
PM = 26.36
DURATION_MONTHS = 8.67
AVG_DEV = round(PM / DURATION_MONTHS, 2)  # ~3.04
DAYS_PER_MONTH = 30.4

# Effort distribution (fractions)
EFFORT = {
    "Management": 0.07,
    "Specification": 0.10,
    "Design": 0.27,
    "Coding": 0.26,
    "Testing": 0.18,
    "Reviews": 0.03,
    "Manuals": 0.09,
}

# Hourly rates (EUR/h)
RATES = {
    "Axel": 40,
    "Bernd": 40,
    "Christine": 45,
    "Diana": 40,
    "Richard": 50,
    "Stefanie": 40,
    "Thomas": 45,
}

# Schedule data
SCHEDULE_CSV = ROOT / "data" / "AMEISE-Schedule-v2-grid.csv"

def load_schedule():
    df = pd.read_csv(SCHEDULE_CSV).fillna("")
    sched = {}
    for _, row in df.iterrows():
        person = str(row["Person"]).strip()
        weeks = {}
        for c in df.columns:
            if c.startswith("W"):
                try:
                    w = int(c[1:])
                except:
                    continue
                val = str(row[c]).strip()
                weeks[w] = "" if val == "nan" else val
        sched[person] = weeks
    return sched

# Helpers to find and write into cells by labels (works with instructor template)
def find_cell(ws, text):
    for r in range(1, min(ws.max_row, 160)+1):
        for c in range(1, min(ws.max_column, 160)+1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() == text:
                return r, c
    return None

def find_cell_startswith(ws, text):
    for r in range(1, min(ws.max_row, 160)+1):
        for c in range(1, min(ws.max_column, 160)+1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().startswith(text):
                return r, c
    return None

def set_right_of_label(ws, label, value):
    pos = find_cell_startswith(ws, label)
    if not pos:
        return False
    r, c = pos
    ws.cell(r, c+1, value)
    return True

def find_week_grid(ws):
    # Find header 'Week' and first week column
    pos = find_cell(ws, "Week")
    if not pos:
        pos = find_cell(ws, "week")
    if not pos:
        return None
    r, c = pos
    # Find numeric '1' to the right
    col1 = None
    for cc in range(c+1, min(c+80, ws.max_column)):
        v = ws.cell(r, cc).value
        if v in (1, "1"):
            col1 = cc
            break
    if not col1:
        col1 = c+1
    return r, col1

def find_person_row(ws, person, start_row=1, max_row=220, search_cols=8):
    for r in range(start_row, min(max_row, ws.max_row)+1):
        for c in range(1, min(search_cols, ws.max_column)+1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() == person:
                return r
    return None

def count_assigned_weeks(ws, person_row, col_week1, n_weeks=40):
    used = 0
    for i in range(n_weeks):
        v = ws.cell(person_row, col_week1 + i).value
        if isinstance(v, str) and v.strip():
            used += 1
    return used

def fill_effort_table(ws):
    # Locate "Effort Distribution"
    pos = find_cell_startswith(ws, "Effort Distribution")
    if not pos:
        return
    r0, c0 = pos
    header_row = r0 + 1
    # Write rows for our categories
    i = 1
    for cat, pct in EFFORT.items():
        r = header_row + i
        ws.cell(r, c0, cat)
        ws.cell(r, c0+1, round(pct*100, 2))           # %
        ws.cell(r, c0+2, round(PM*pct, 4))            # PM
        months = DURATION_MONTHS * pct
        ws.cell(r, c0+3, round(months, 4))            # months
        ws.cell(r, c0+4, round(months*DAYS_PER_MONTH, 2))  # days
        i += 1
    # Totals
    ws.cell(header_row + len(EFFORT) + 1, c0, "Total Effort")
    ws.cell(header_row + len(EFFORT) + 1, c0+2, round(PM, 2))
    ws.cell(header_row + len(EFFORT) + 1, c0+3, round(DURATION_MONTHS, 2))
    ws.cell(header_row + len(EFFORT) + 1, c0+4, round(DURATION_MONTHS*DAYS_PER_MONTH, 0))

def fill_schedule_grid(ws, sched):
    info = find_week_grid(ws)
    if not info:
        return
    header_row, col_week1 = info
    start_row = header_row + 1
    n_weeks = 40
    for person, weeks in sched.items():
        r = find_person_row(ws, person, start_row=start_row)
        if not r:
            continue
        for i in range(1, n_weeks+1):
            token = weeks.get(i, "")
            ws.cell(r, col_week1 + (i-1), token if token else None)

def fill_cost_table(ws):
    pos = find_cell_startswith(ws, "Cost Estimation")
    if not pos:
        return
    r0, c0 = pos
    name_col = c0
    rate_col = c0 + 1
    weeks_col = c0 + 2
    months_col = c0 + 3
    hours_col = c0 + 4
    total_col = c0 + 5

    # Get week grid to count activity cells per person
    info = find_week_grid(ws)
    if not info:
        return
    _, col_week1 = info

    # Iterate rows below header
    row = r0 + 2
    grand_total = 0.0
    while row <= ws.max_row:
        name = ws.cell(row, name_col).value
        if not name or (isinstance(name, str) and name.strip() == ""):
            break
        name = str(name).strip()
        if name not in RATES:
            row += 1
            continue
        rate = RATES[name]
        person_row = find_person_row(ws, name, start_row=1)
        weeks_used = count_assigned_weeks(ws, person_row, col_week1) if person_row else 0
        hours = weeks_used * 40
        months = round((weeks_used * 7.0) / DAYS_PER_MONTH, 2)
        total = round(hours * rate, 2)

        ws.cell(row, rate_col, rate)
        ws.cell(row, weeks_col, weeks_used)
        ws.cell(row, months_col, months)
        ws.cell(row, hours_col, hours)
        ws.cell(row, total_col, total)

        grand_total += total
        row += 1

    # Find "Total Project Costs" row and write total
    for rr in range(row, row+10):
        v = ws.cell(rr, name_col).value
        if isinstance(v, str) and v.strip().startswith("Total Project Costs"):
            ws.cell(rr, total_col, round(grand_total, 2))
            break

def main():
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
    sched = load_schedule()
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Top-left metrics (right to labels)
    set_right_of_label(ws, "Est. Proj. Duration", round(DURATION_MONTHS, 2))
    set_right_of_label(ws, "Est. Avg. Developers", AVG_DEV)
    set_right_of_label(ws, "Est. PersonMonths", round(PM, 2))

    fill_effort_table(ws)
    fill_schedule_grid(ws, sched)
    fill_cost_table(ws)

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Saved: {OUT_PATH}")

# Reused helper
def set_right_of_label(ws, label, value):
    pos = find_cell_startswith(ws, label)
    if not pos:
        return False
    r, c = pos
    ws.cell(r, c+1, value)
    return True

if __name__ == "__main__":
    main()
