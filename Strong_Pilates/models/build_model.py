"""
Strong Pilates Kingston upon Thames — Financial Model Builder
================================================================

Builds an Excel workbook with:
  - Assumptions sheet (every input)
  - 10-year monthly P&L for 3 scenarios (Optimistic / Base / Pessimistic)
  - Annual roll-up
  - Cumulative cash & equity returns
  - Scenario summary at 2y / 3y / 5y / 10y for the friend's £250k

Run:  python3 build_model.py
Output: ../report/Strong_Pilates_Kingston_Model.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule


# ---------------------------------------------------------------------------
# 1. ASSUMPTIONS
# ---------------------------------------------------------------------------

SCENARIOS = {
    "Optimistic": {
        # capacity & ramp
        "beds": 18,
        "classes_per_week": 56,
        "year1_util_pct": 0.38,
        "year2_util_pct": 0.55,
        "year3_util_pct": 0.65,
        "year5_util_pct": 0.68,
        "ramp_months": 9,         # months from open to "year-1 average" utilisation
        # pricing
        "rev_per_filled_seat": 21.0,
        "annual_price_growth": 0.03,
        # capex
        "build_out": 200_000,
        "equipment": 165_000,
        "initial_franchise_fee": 45_000,
        "preopen_marketing": 35_000,
        "working_capital_reserve": 50_000,
        # opex
        "annual_rent_all_in": 88_000,    # rent + rates + service + insurance
        "annual_instructor_cost": 110_000,
        "annual_foh_cost": 38_000,
        "owner_salary_y1": 0,
        "owner_salary_y2plus": 35_000,
        "annual_other_opex": 65_000,
        "year1_marketing": 45_000,
        "steady_marketing_pct": 0.06,    # of revenue
        # franchise fees
        "royalty_pct": 0.08,
        "brand_fund_pct": 0.02,
        # debt
        "debt_apr": 0.085,
        "debt_term_years": 7,
    },
    "Base": {
        "beds": 16,
        "classes_per_week": 49,
        "year1_util_pct": 0.30,
        "year2_util_pct": 0.46,
        "year3_util_pct": 0.55,
        "year5_util_pct": 0.58,
        "ramp_months": 12,
        "rev_per_filled_seat": 18.0,
        "annual_price_growth": 0.025,
        "build_out": 230_000,
        "equipment": 170_000,
        "initial_franchise_fee": 50_000,
        "preopen_marketing": 50_000,
        "working_capital_reserve": 60_000,
        "annual_rent_all_in": 95_000,
        "annual_instructor_cost": 118_000,
        "annual_foh_cost": 42_000,
        "owner_salary_y1": 0,
        "owner_salary_y2plus": 35_000,
        "annual_other_opex": 75_000,
        "year1_marketing": 55_000,
        "steady_marketing_pct": 0.07,
        "royalty_pct": 0.08,
        "brand_fund_pct": 0.02,
        "debt_apr": 0.10,
        "debt_term_years": 7,
    },
    "Pessimistic": {
        "beds": 16,
        "classes_per_week": 49,
        "year1_util_pct": 0.22,
        "year2_util_pct": 0.34,
        "year3_util_pct": 0.42,
        "year5_util_pct": 0.45,
        "ramp_months": 18,
        "rev_per_filled_seat": 16.0,
        "annual_price_growth": 0.02,
        "build_out": 270_000,
        "equipment": 180_000,
        "initial_franchise_fee": 55_000,
        "preopen_marketing": 60_000,
        "working_capital_reserve": 70_000,
        "annual_rent_all_in": 105_000,
        "annual_instructor_cost": 125_000,
        "annual_foh_cost": 45_000,
        "owner_salary_y1": 0,
        "owner_salary_y2plus": 30_000,
        "annual_other_opex": 85_000,
        "year1_marketing": 70_000,
        "steady_marketing_pct": 0.08,
        "royalty_pct": 0.08,
        "brand_fund_pct": 0.02,
        "debt_apr": 0.115,
        "debt_term_years": 7,
    },
}

EQUITY = 250_000   # the friend's investment
COST_INFLATION = 0.03  # opex grows 3%/year


# ---------------------------------------------------------------------------
# 2. MODEL ENGINE
# ---------------------------------------------------------------------------

def utilisation_curve(scn, month):
    """
    Returns the seat utilisation for a given month index (0-indexed, 0 = first month open).
    Uses a linear ramp from a starting low to year-1 target over `ramp_months`,
    then linearly transitions year1->year2->year3->year5 targets,
    holding year-5 from then on.
    """
    ramp = scn["ramp_months"]
    y1, y2, y3, y5 = (scn["year1_util_pct"], scn["year2_util_pct"],
                      scn["year3_util_pct"], scn["year5_util_pct"])

    # initial utilisation at open: 1/3 of y1
    start = y1 / 3.0

    if month < ramp:
        return start + (y1 - start) * (month / ramp)

    if month < 12:        # rest of year 1
        return y1
    if month < 24:        # year 2
        # linear from y1 -> y2 across year 2
        frac = (month - 12) / 12.0
        return y1 + (y2 - y1) * frac
    if month < 36:        # year 3
        frac = (month - 24) / 12.0
        return y2 + (y3 - y2) * frac
    if month < 60:        # year 4-5: linear ramp y3 -> y5
        frac = (month - 36) / 24.0
        return y3 + (y5 - y3) * frac
    # year 6+: hold year 5
    return y5


def build_monthly(scn):
    """Returns a list of dicts, one per month, for 120 months (10 years)."""
    rows = []
    months = 120

    annual_seats = scn["beds"] * scn["classes_per_week"] * 52  # full annual seat capacity
    monthly_seat_capacity = annual_seats / 12.0

    # debt sizing — base case derived later, but per-scenario debt
    capex_total = (scn["build_out"] + scn["equipment"]
                   + scn["initial_franchise_fee"] + scn["preopen_marketing"]
                   + scn["working_capital_reserve"])
    debt = max(0.0, capex_total - EQUITY)

    # amortising loan monthly payment
    apr = scn["debt_apr"]
    n = scn["debt_term_years"] * 12
    if debt > 0:
        r = apr / 12.0
        if r > 0:
            monthly_debt_payment = debt * r * (1 + r) ** n / ((1 + r) ** n - 1)
        else:
            monthly_debt_payment = debt / n
    else:
        monthly_debt_payment = 0.0

    debt_balance = debt

    for m in range(months):
        year_index = m // 12  # 0..9
        # cost inflator
        cost_inflator = (1 + COST_INFLATION) ** year_index
        price_inflator = (1 + scn["annual_price_growth"]) ** year_index

        util = utilisation_curve(scn, m)
        seats_sold = monthly_seat_capacity * util
        rev_per_seat = scn["rev_per_filled_seat"] * price_inflator
        revenue = seats_sold * rev_per_seat

        # opex
        rent = scn["annual_rent_all_in"] / 12 * cost_inflator
        instructors = scn["annual_instructor_cost"] / 12 * cost_inflator
        foh = scn["annual_foh_cost"] / 12 * cost_inflator
        owner = (scn["owner_salary_y1"] if year_index == 0 else scn["owner_salary_y2plus"]) / 12
        other_opex = scn["annual_other_opex"] / 12 * cost_inflator

        # marketing: heavy year 1 then % of revenue
        if year_index == 0:
            marketing = scn["year1_marketing"] / 12
        else:
            marketing = revenue * scn["steady_marketing_pct"]

        royalty = revenue * scn["royalty_pct"]
        brand_fund = revenue * scn["brand_fund_pct"]

        total_opex = (rent + instructors + foh + owner + other_opex
                      + marketing + royalty + brand_fund)

        ebitda = revenue - total_opex

        # debt service: split into interest and principal for clarity
        if debt_balance > 0:
            interest = debt_balance * (apr / 12.0)
            principal = max(0.0, monthly_debt_payment - interest)
            principal = min(principal, debt_balance)
            debt_balance -= principal
        else:
            interest = principal = 0.0

        # depreciation: build-out 10y SL; equipment 5y SL
        dep_buildout = scn["build_out"] / (10 * 12)
        dep_equipment = scn["equipment"] / (5 * 12) if year_index < 5 else 0
        depreciation = dep_buildout + dep_equipment

        ebit = ebitda - depreciation
        pbt = ebit - interest

        # Tax: 25% on positive PBT only (UK CT main rate post-2023)
        tax = max(0.0, pbt) * 0.25
        pat = pbt - tax

        # cash flow to equity = EBITDA - interest - principal - tax
        cash_to_equity = ebitda - interest - principal - tax

        rows.append({
            "month": m + 1,
            "year": year_index + 1,
            "utilisation": util,
            "seats_sold": seats_sold,
            "rev_per_seat": rev_per_seat,
            "revenue": revenue,
            "rent": rent,
            "instructors": instructors,
            "foh": foh,
            "owner": owner,
            "other_opex": other_opex,
            "marketing": marketing,
            "royalty": royalty,
            "brand_fund": brand_fund,
            "total_opex": total_opex,
            "ebitda": ebitda,
            "depreciation": depreciation,
            "ebit": ebit,
            "interest": interest,
            "principal": principal,
            "pbt": pbt,
            "tax": tax,
            "pat": pat,
            "cash_to_equity": cash_to_equity,
            "debt_balance_eom": debt_balance,
        })

    return rows, debt, monthly_debt_payment


def annual_rollup(rows):
    out = {}
    for r in rows:
        y = r["year"]
        if y not in out:
            out[y] = {k: 0.0 for k in r if k not in ("month", "year", "utilisation",
                                                      "rev_per_seat", "debt_balance_eom")}
            out[y]["months"] = 0
            out[y]["avg_util"] = 0.0
        out[y]["months"] += 1
        out[y]["avg_util"] += r["utilisation"]
        for k in out[y]:
            if k in ("months", "avg_util"):
                continue
            out[y][k] += r[k]
    for y in out:
        out[y]["avg_util"] /= out[y]["months"]
    return out


# ---------------------------------------------------------------------------
# 3. EXCEL OUTPUT
# ---------------------------------------------------------------------------

THIN = Side(border_style="thin", color="DDDDDD")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SECTION_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
NEG_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")


def fmt_money(cell):
    cell.number_format = '£#,##0;[Red]-£#,##0'


def fmt_pct(cell):
    cell.number_format = '0.0%'


def write_assumptions_sheet(wb):
    ws = wb.create_sheet("Assumptions")
    ws["A1"] = "Strong Pilates Kingston — Assumptions"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:E1")

    ws["A3"] = "Friend's equity"
    ws["B3"] = EQUITY
    fmt_money(ws["B3"])
    ws["A4"] = "Cost inflation p.a."
    ws["B4"] = COST_INFLATION
    fmt_pct(ws["B4"])

    headers = ["Parameter", "Optimistic", "Base", "Pessimistic"]
    row = 6
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    row += 1
    keys = list(SCENARIOS["Base"].keys())
    pct_keys = {"year1_util_pct", "year2_util_pct", "year3_util_pct", "year5_util_pct",
                "annual_price_growth", "steady_marketing_pct", "royalty_pct",
                "brand_fund_pct", "debt_apr"}
    for k in keys:
        ws.cell(row=row, column=1, value=k)
        for i, sc in enumerate(["Optimistic", "Base", "Pessimistic"], 2):
            v = SCENARIOS[sc][k]
            cell = ws.cell(row=row, column=i, value=v)
            if k in pct_keys:
                fmt_pct(cell)
            elif isinstance(v, (int, float)) and abs(v) > 100:
                fmt_money(cell)
        row += 1

    # column widths
    ws.column_dimensions["A"].width = 32
    for col in "BCDE":
        ws.column_dimensions[col].width = 16

    # totals
    row += 1
    ws.cell(row=row, column=1, value="Total project cost").font = Font(bold=True)
    for i, sc in enumerate(["Optimistic", "Base", "Pessimistic"], 2):
        s = SCENARIOS[sc]
        total = (s["build_out"] + s["equipment"] + s["initial_franchise_fee"]
                 + s["preopen_marketing"] + s["working_capital_reserve"])
        c = ws.cell(row=row, column=i, value=total)
        fmt_money(c)
        c.font = Font(bold=True)
    row += 1
    ws.cell(row=row, column=1, value="Equity (friend's £250k)").font = Font(bold=True)
    for i in range(2, 5):
        c = ws.cell(row=row, column=i, value=EQUITY)
        fmt_money(c)
    row += 1
    ws.cell(row=row, column=1, value="Required debt").font = Font(bold=True)
    for i, sc in enumerate(["Optimistic", "Base", "Pessimistic"], 2):
        s = SCENARIOS[sc]
        total = (s["build_out"] + s["equipment"] + s["initial_franchise_fee"]
                 + s["preopen_marketing"] + s["working_capital_reserve"])
        c = ws.cell(row=row, column=i, value=max(0, total - EQUITY))
        fmt_money(c)
        c.font = Font(bold=True)


def write_monthly_sheet(wb, name, rows):
    ws = wb.create_sheet(name)
    headers = ["Month", "Year", "Util %", "Seats sold", "Rev/seat",
               "Revenue", "Rent+Rates", "Instructors", "FOH", "Owner",
               "Other opex", "Marketing", "Royalty 8%", "Brand fund 2%",
               "Total opex", "EBITDA", "Depreciation", "EBIT",
               "Interest", "Principal", "PBT", "Tax", "PAT",
               "Cash to equity", "Debt balance"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, r in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=r["month"])
        ws.cell(row=i, column=2, value=r["year"])
        u = ws.cell(row=i, column=3, value=r["utilisation"]); fmt_pct(u)
        ws.cell(row=i, column=4, value=round(r["seats_sold"], 0))
        ws.cell(row=i, column=5, value=round(r["rev_per_seat"], 2))
        ws.cell(row=i, column=5).number_format = '£#,##0.00'
        col_keys = ["revenue", "rent", "instructors", "foh", "owner",
                    "other_opex", "marketing", "royalty", "brand_fund",
                    "total_opex", "ebitda", "depreciation", "ebit",
                    "interest", "principal", "pbt", "tax", "pat",
                    "cash_to_equity", "debt_balance_eom"]
        for j, k in enumerate(col_keys, 6):
            c = ws.cell(row=i, column=j, value=round(r[k], 0))
            fmt_money(c)
        # highlight negative EBITDA rows
        if r["ebitda"] < 0:
            ws.cell(row=i, column=16).fill = NEG_FILL
    ws.freeze_panes = "C2"
    for col_letter in "ABCDEFGHIJKLMNOPQRSTUVWXY":
        ws.column_dimensions[col_letter].width = 13
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 7


def write_annual_sheet(wb, ann_by_scenario):
    ws = wb.create_sheet("Annual P&L")
    ws["A1"] = "Annual P&L by scenario"
    ws["A1"].font = Font(bold=True, size=14)

    line_keys = [
        ("Average utilisation %", "avg_util", "pct"),
        ("Revenue", "revenue", "money"),
        ("Rent + rates + service", "rent", "money"),
        ("Instructor cost", "instructors", "money"),
        ("Front of house", "foh", "money"),
        ("Owner salary", "owner", "money"),
        ("Other opex", "other_opex", "money"),
        ("Marketing", "marketing", "money"),
        ("Royalty 8%", "royalty", "money"),
        ("Brand fund 2%", "brand_fund", "money"),
        ("TOTAL OPEX", "total_opex", "money"),
        ("EBITDA", "ebitda", "money"),
        ("Depreciation", "depreciation", "money"),
        ("EBIT", "ebit", "money"),
        ("Interest", "interest", "money"),
        ("PBT", "pbt", "money"),
        ("Tax", "tax", "money"),
        ("PAT", "pat", "money"),
        ("Cash to equity (after debt service & tax)", "cash_to_equity", "money"),
    ]

    row = 3
    for sc_name, ann in ann_by_scenario.items():
        ws.cell(row=row, column=1, value=sc_name).font = Font(bold=True, size=12)
        ws.cell(row=row, column=1).fill = SECTION_FILL
        row += 1
        # year header
        ws.cell(row=row, column=1, value="Line").font = HEADER_FONT
        ws.cell(row=row, column=1).fill = HEADER_FILL
        for y in range(1, 11):
            c = ws.cell(row=row, column=y + 1, value=f"Year {y}")
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
            c.alignment = Alignment(horizontal="center")
        row += 1
        for label, key, kind in line_keys:
            ws.cell(row=row, column=1, value=label)
            for y in range(1, 11):
                v = ann[y][key]
                c = ws.cell(row=row, column=y + 1, value=round(v, 0) if kind == "money" else v)
                if kind == "money":
                    fmt_money(c)
                else:
                    fmt_pct(c)
                if kind == "money" and v < 0:
                    c.fill = NEG_FILL
            row += 1
        row += 2

    ws.column_dimensions["A"].width = 38
    for col in range(2, 12):
        ws.column_dimensions[get_column_letter(col)].width = 13


def write_summary_sheet(wb, scenarios_data):
    ws = wb.create_sheet("Summary", 0)
    ws["A1"] = "Strong Pilates Kingston — Investment Summary"
    ws["A1"].font = Font(bold=True, size=16)
    ws.merge_cells("A1:F1")

    ws["A3"] = "Friend's equity at risk"
    ws["B3"] = EQUITY
    fmt_money(ws["B3"])
    ws["B3"].font = Font(bold=True, size=12)

    headers = ["Metric", "Optimistic", "Base", "Pessimistic"]
    row = 5
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    metric_blocks = [
        ("Project cost", "project_cost", "money"),
        ("Required debt", "debt", "money"),
        ("Year-1 revenue", "y1_rev", "money"),
        ("Year-1 EBITDA", "y1_ebitda", "money"),
        ("Year-1 cash to equity", "y1_cte", "money"),
        ("Year-2 revenue", "y2_rev", "money"),
        ("Year-2 EBITDA", "y2_ebitda", "money"),
        ("Year-2 cash to equity", "y2_cte", "money"),
        ("Year-3 revenue", "y3_rev", "money"),
        ("Year-3 EBITDA", "y3_ebitda", "money"),
        ("Year-3 cash to equity", "y3_cte", "money"),
        ("Year-5 revenue", "y5_rev", "money"),
        ("Year-5 EBITDA", "y5_ebitda", "money"),
        ("Year-5 cash to equity", "y5_cte", "money"),
        ("Year-10 revenue", "y10_rev", "money"),
        ("Year-10 EBITDA", "y10_ebitda", "money"),
        ("Year-10 cash to equity", "y10_cte", "money"),
        ("Cumulative cash to equity Y2", "cum2", "money"),
        ("Cumulative cash to equity Y3", "cum3", "money"),
        ("Cumulative cash to equity Y5", "cum5", "money"),
        ("Cumulative cash to equity Y10", "cum10", "money"),
        ("Equity payback Y2 (cum/equity)", "payback2", "pct"),
        ("Equity payback Y3", "payback3", "pct"),
        ("Equity payback Y5", "payback5", "pct"),
        ("Equity payback Y10", "payback10", "pct"),
        ("Months to operating breakeven", "months_to_op_be", "num"),
    ]

    row = 6
    for label, k, kind in metric_blocks:
        ws.cell(row=row, column=1, value=label)
        for col, sc in enumerate(["Optimistic", "Base", "Pessimistic"], 2):
            v = scenarios_data[sc][k]
            c = ws.cell(row=row, column=col, value=v if v is not None else "—")
            if kind == "money":
                fmt_money(c)
                if isinstance(v, (int, float)) and v < 0:
                    c.fill = NEG_FILL
            elif kind == "pct":
                fmt_pct(c)
            else:
                c.number_format = '0'
        row += 1

    ws.column_dimensions["A"].width = 42
    for col in "BCD":
        ws.column_dimensions[col].width = 16

    # narrative box
    row += 2
    ws.cell(row=row, column=1, value="How to read this summary").font = Font(bold=True, size=12)
    row += 1
    notes = [
        "1. 'Cash to equity' = EBITDA – interest – principal – tax. It is the cash the studio",
        "    can pay out to its owner each year. Negative numbers mean the studio is losing money",
        "    AND requires the owner to inject more cash to keep paying the bills.",
        "2. 'Equity payback' shows the cumulative cash returned to the owner as a % of the £250k.",
        "    A figure under 100% by year 10 means the owner has not yet made their money back —",
        "    not even ignoring the time value of money.",
        "3. 'Months to operating breakeven' = first month where monthly EBITDA > 0 and stays > 0.",
        "    A '—' means it never breaks even within the 10-year window.",
        "4. The Pessimistic scenario is NOT a worst-case. A real worst-case is the studio failing",
        "    in year 1 or 2 and the owner losing the entire £250k AND owing the bank the loan",
        "    AND being on the hook for personal guarantees on the lease.",
    ]
    for line in notes:
        ws.cell(row=row, column=1, value=line)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1


def compute_summary(rows, scn):
    ann = annual_rollup(rows)

    capex = (scn["build_out"] + scn["equipment"] + scn["initial_franchise_fee"]
             + scn["preopen_marketing"] + scn["working_capital_reserve"])
    debt = max(0, capex - EQUITY)

    cum = 0
    cum_by_year = {}
    for y in range(1, 11):
        cum += ann[y]["cash_to_equity"]
        cum_by_year[y] = cum

    # months to operating breakeven (first month with positive EBITDA that stays positive)
    months_to_be = None
    for i, r in enumerate(rows):
        if r["ebitda"] > 0:
            # check that >75% of subsequent 6 months are also positive
            window = rows[i:i + 6]
            pos = sum(1 for w in window if w["ebitda"] > 0)
            if pos >= 4:
                months_to_be = i + 1
                break

    return {
        "project_cost": capex,
        "debt": debt,
        "y1_rev": ann[1]["revenue"],
        "y1_ebitda": ann[1]["ebitda"],
        "y1_cte": ann[1]["cash_to_equity"],
        "y2_rev": ann[2]["revenue"],
        "y2_ebitda": ann[2]["ebitda"],
        "y2_cte": ann[2]["cash_to_equity"],
        "y3_rev": ann[3]["revenue"],
        "y3_ebitda": ann[3]["ebitda"],
        "y3_cte": ann[3]["cash_to_equity"],
        "y5_rev": ann[5]["revenue"],
        "y5_ebitda": ann[5]["ebitda"],
        "y5_cte": ann[5]["cash_to_equity"],
        "y10_rev": ann[10]["revenue"],
        "y10_ebitda": ann[10]["ebitda"],
        "y10_cte": ann[10]["cash_to_equity"],
        "cum2": cum_by_year[2],
        "cum3": cum_by_year[3],
        "cum5": cum_by_year[5],
        "cum10": cum_by_year[10],
        "payback2": cum_by_year[2] / EQUITY,
        "payback3": cum_by_year[3] / EQUITY,
        "payback5": cum_by_year[5] / EQUITY,
        "payback10": cum_by_year[10] / EQUITY,
        "months_to_op_be": months_to_be if months_to_be else "Never (10y)",
    }


def main():
    wb = Workbook()
    wb.remove(wb.active)

    write_assumptions_sheet(wb)

    ann_by_scenario = {}
    summary_by_scenario = {}
    for sc_name, scn in SCENARIOS.items():
        rows, debt, _ = build_monthly(scn)
        ann_by_scenario[sc_name] = annual_rollup(rows)
        summary_by_scenario[sc_name] = compute_summary(rows, scn)
        write_monthly_sheet(wb, f"Monthly_{sc_name}", rows)

    write_annual_sheet(wb, ann_by_scenario)
    write_summary_sheet(wb, summary_by_scenario)

    out_path = "../report/Strong_Pilates_Kingston_Model.xlsx"
    wb.save(out_path)
    print(f"Wrote {out_path}")

    # also dump the summary to text for the report
    print("\n=== SUMMARY ===")
    for sc, s in summary_by_scenario.items():
        print(f"\n--- {sc} ---")
        for k, v in s.items():
            if isinstance(v, float):
                print(f"  {k}: £{v:,.0f}" if abs(v) > 1 else f"  {k}: {v:.1%}")
            else:
                print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
