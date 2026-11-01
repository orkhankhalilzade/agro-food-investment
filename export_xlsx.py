"""Excel export (new workbook; the original company templates are never touched).

Sheets
  Operations          input values + live Excel formulas (same logic as the company templates)
  Operations (values) the same rows with the values calculated by the app (calc.py)
  Check               compares the two sheets cell by cell ("OK" / "DIFF") -- open in Excel to verify
  Normatives, Machinery inventory, Implement inventory, Notes

All formulas of one row refer only to cells of the same row, so rows can be copied into a
combined workbook without breaking anything.
"""
from __future__ import annotations

import datetime as dt
import io
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

import core
from calc import Norms, calculate

# (key, header exactly as in the requested column list, kind)
#   kind: "in" = input, "calc" = calculated, "norm" = normative copied onto the row
COLUMNS: list[tuple[str, str, str]] = [
    ("company", "Company", "in"), ("area_ha", "Area(ha)", "in"), ("orchard", "Orchard Name", "in"),
    ("stage", "Phenological Status", "in"), ("month", "Month", "in"), ("week", "Week", "in"),
    ("pests", "Pests", "in"), ("diseases", "Diseases", "in"),
    ("nutrition", "Macro- and Micronutrient Feeding", "in"),
    ("regulators", "Plant Regulators (acids, hormones, sun protection)", "in"),
    ("adjuvant", "Adjuvant / pH Corrector", "in"), ("herbicide", "Herbicide", "in"),
    ("pt", "operations", "in"), ("solution_per_ha", "Spraying Solution per ha", "in"),
    ("row_length", "Average row length, m", "in"), ("rows", "Rows", "in"),
    ("max_duration", "Max duration operation time, days", "in"), ("tractor", "Tractor", "in"),
    ("speed", "Speed Tractor, km/h (for current operation)", "in"),
    ("fuel_with_sprayer", "Tractor, fuel consumption with sprayer, l/h", "in"),
    ("oil_with_sprayer", "Oil consumption with sprayer, l/h", "in"),
    ("fuel_without_sprayer", "Tractor, fuel consumption without sprayer, l/h (transit/reload)", "in"),
    ("oil_without_sprayer", "Oil consumption without sprayer, l/h (transit/reload)", "in"),
    ("maintenance_time_h", "time for machinery maintenance", "in"),
    ("effective_shift_h", "Work shift, hours", "calc"),
    ("tank_capacity_l", "Sprayer, tank liter (Agrotek)", "in"),
    ("total_solution_l", "Total solution per field", "calc"),
    ("tanks", "Tanks per Area, unit", "calc"),
    ("productivity_ha_h", "Productivity, ha/hour (sprayer+tractor working)", "in"),
    ("reload_minutes", "Time for tank reloading, min", "in"),
    ("reload_total_h", "Total Time for tank reloading, hour/operation", "calc"),
    ("reload_per_ha", "Reloading per Ha", "calc"),
    ("pure_productivity", "Pure working productivity", "calc"),
    ("output_per_shift_ha", "Output per shift per 1 tractor, ha", "calc"),
    ("time_h", "Time, hours/operation", "calc"),
    ("engine_hours", "Engine hours (including turns)", "calc"),
    ("working_days", "Working days", "calc"),
    ("diesel_active_l", "Disel consumption, total, liter (activ)", "calc"),
    ("lubricant_active_l", "Lubricants (activ)", "calc"),
    ("diesel_transit_l", "Disel consumption, total, liter (transit)", "calc"),
    ("lubricant_transit_l", "Lubricants (transit), liter", "calc"),
    ("diesel_cost", "Cost of diesel, azn/liter", "calc"),
    ("lubricant_cost", "Cost of lubricants, azn/liter", "calc"),
    ("total_fuel_lubricants", "Total fuel and lubricants", "calc"),
    ("salary_per_hour", "Salary, gross per hour", "in"),
    ("salary_gross", "Salary, gross", "calc"),
    ("total_cost", "Total cost", "calc"),
    # extra columns (needed so that every row is self-contained)
    ("tractors", "Number of tractors", "in"),
    ("n_work_shift_h", "Norm: Work shift, hours", "norm"),
    ("n_round_up_tanks", "Norm: Tanks per Area (1 = round up, 0 = fraction)", "norm"),
    ("n_engine_factor", "Norm: Engine hours (including turns) factor", "norm"),
    ("n_diesel_cost", "Norm: Cost of diesel, AZN/L", "norm"),
    ("n_lubricant_cost", "Norm: Cost of lubricants, AZN/L", "norm"),
    ("status", "Status", "info"),
]
KEYS = [c[0] for c in COLUMNS]
COL = {k: L(i + 1) for i, k in enumerate(KEYS)}
CALC_KEYS = [k for k, _, kind in COLUMNS if kind == "calc"]

GREEN, MID, LIGHT = "144927", "5F8F54", "DAEFDA"
HEAD_FONT = Font(bold=True, color="F5FAF4")


def _formulas(r: int) -> dict[str, str]:
    """Excel formulas of row r; same logic and guards as the company templates (see calc.py)."""
    c = {k: f"{COL[k]}{r}" for k in KEYS}
    return {
        "effective_shift_h": f'=IF(ISNUMBER({c["maintenance_time_h"]}),{c["n_work_shift_h"]}-{c["maintenance_time_h"]},"")',
        "total_solution_l": f'=IF(AND(ISNUMBER({c["solution_per_ha"]}),ISNUMBER({c["area_ha"]})),{c["solution_per_ha"]}*{c["area_ha"]},"")',
        "tanks": (f'=IF(AND(ISNUMBER({c["total_solution_l"]}),ISNUMBER({c["tank_capacity_l"]})),'
                  f'IF({c["tank_capacity_l"]}>0,IF({c["n_round_up_tanks"]}=1,ROUNDUP({c["total_solution_l"]}/{c["tank_capacity_l"]},0),'
                  f'{c["total_solution_l"]}/{c["tank_capacity_l"]}),""),"")'),
        "reload_total_h": f'=IF(AND(ISNUMBER({c["tanks"]}),ISNUMBER({c["reload_minutes"]})),{c["reload_minutes"]}*{c["tanks"]}/60,"")',
        "reload_per_ha": f'=IF(ISNUMBER({c["reload_total_h"]}),{c["reload_total_h"]}/{c["area_ha"]},"")',
        "time_h": (f'=IF(AND(ISNUMBER({c["productivity_ha_h"]}),ISNUMBER({c["reload_total_h"]})),'
                   f'IF({c["productivity_ha_h"]}>0,{c["area_ha"]}/{c["productivity_ha_h"]}+{c["reload_total_h"]},""),"")'),
        "pure_productivity": f'=IF(ISNUMBER({c["time_h"]}),IF({c["time_h"]}>0,{c["area_ha"]}/{c["time_h"]},""),"")',
        "output_per_shift_ha": f'=IF(AND(ISNUMBER({c["pure_productivity"]}),ISNUMBER({c["effective_shift_h"]})),{c["pure_productivity"]}*{c["effective_shift_h"]},"")',
        "engine_hours": f'=IF(ISNUMBER({c["time_h"]}),({c["time_h"]}-{c["reload_total_h"]})+({c["reload_total_h"]}*{c["n_engine_factor"]}),"")',
        "working_days": (f'=IF(AND(ISNUMBER({c["time_h"]}),ISNUMBER({c["effective_shift_h"]})),'
                         f'IF({c["effective_shift_h"]}>0,{c["time_h"]}/({c["effective_shift_h"]}*MAX(1,N({c["tractors"]}))),""),"")'),
        "diesel_active_l": f'=IF(AND(ISNUMBER({c["time_h"]}),ISNUMBER({c["fuel_with_sprayer"]})),({c["time_h"]}-{c["reload_total_h"]})*{c["fuel_with_sprayer"]},"")',
        "lubricant_active_l": f'=IF(AND(ISNUMBER({c["time_h"]}),ISNUMBER({c["oil_with_sprayer"]})),({c["time_h"]}-{c["reload_total_h"]})*{c["oil_with_sprayer"]},"")',
        "diesel_transit_l": f'=IF(AND(ISNUMBER({c["reload_total_h"]}),ISNUMBER({c["fuel_without_sprayer"]})),{c["reload_total_h"]}*{c["fuel_without_sprayer"]},"")',
        "lubricant_transit_l": f'=IF(AND(ISNUMBER({c["reload_total_h"]}),ISNUMBER({c["oil_without_sprayer"]})),{c["reload_total_h"]}*{c["oil_without_sprayer"]},"")',
        "diesel_cost": f'=IF(ISNUMBER({c["diesel_active_l"]}),{c["n_diesel_cost"]}*({c["diesel_active_l"]}+N({c["diesel_transit_l"]})),"")',
        "lubricant_cost": f'=IF(ISNUMBER({c["lubricant_active_l"]}),{c["n_lubricant_cost"]}*({c["lubricant_active_l"]}+N({c["lubricant_transit_l"]})),"")',
        "total_fuel_lubricants": f'=IF(OR(ISNUMBER({c["diesel_cost"]}),ISNUMBER({c["lubricant_cost"]})),N({c["diesel_cost"]})+N({c["lubricant_cost"]}),"")',
        "salary_gross": f'=IF(AND(ISNUMBER({c["salary_per_hour"]}),ISNUMBER({c["time_h"]})),{c["salary_per_hour"]}*{c["time_h"]},"")',
        "total_cost": f'=IF(AND(ISNUMBER({c["salary_gross"]}),ISNUMBER({c["total_fuel_lubricants"]})),{c["salary_gross"]}+{c["total_fuel_lubricants"]},"")',
    }


def _input_values(cfg: core.Config, op: dict) -> dict:
    i, n = op["inputs"], op["norms"]
    v = {"company": op["company"], "area_ha": cfg.area(op["company"], op["orchard"]),
         "orchard": op["orchard"] or "None", "stage": op["stage"], "month": op["month"], "week": op["week"],
         "pt": op["pt"], "tractor": i.get("tractor") or None, "status": None,
         "n_work_shift_h": n["work_shift_h"], "n_round_up_tanks": n["round_up_tanks"],
         "n_engine_factor": n["engine_factor"], "n_diesel_cost": n["diesel_cost"],
         "n_lubricant_cost": n["lubricant_cost"]}
    for cat in core.CATEGORIES:
        v[cat] = op["products"].get(cat) or None
    for k in core.NUMERIC_INPUTS:
        v[k] = i.get(k)
    return v


def _style_header(ws, kinds: list[str]):
    for idx, kind in enumerate(kinds, start=1):
        cell = ws.cell(1, idx) if ws.title.startswith("Operations") else None
        if cell is None:
            continue
        cell.font = HEAD_FONT
        cell.fill = PatternFill("solid", fgColor=MID if kind == "calc" else GREEN)
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def _plain_header(ws, row: int, ncols: int):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = HEAD_FONT
        cell.fill = PatternFill("solid", fgColor=GREEN)
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def build_workbook(cfg: core.Config, company: str, operations: list[dict], norms: dict[str, dict],
                   inventory: dict[str, dict]) -> bytes:
    ops = sorted((o for o in operations if o["company"] == company), key=core.sort_key)
    wb = Workbook()
    ws_f = wb.active
    ws_f.title = "Operations"
    ws_v = wb.create_sheet("Operations (values)")
    ws_c = wb.create_sheet("Check")
    ws_n = wb.create_sheet("Normatives")
    ws_t = wb.create_sheet("Machinery inventory")
    ws_i = wb.create_sheet("Implement inventory")
    ws_x = wb.create_sheet("Notes")

    headers = [h for _, h, _ in COLUMNS]
    kinds = [k for _, _, k in COLUMNS]
    for ws in (ws_f, ws_v):
        for idx, h in enumerate(headers, start=1):
            ws.cell(1, idx, h)
        _style_header(ws, kinds)
        ws.freeze_panes = "D2"
        ws.row_dimensions[1].height = 48
        for idx, key in enumerate(KEYS, start=1):
            ws.column_dimensions[L(idx)].width = 34 if key in core.CATEGORIES else 16

    for n, op in enumerate(ops):
        r = n + 2
        vals = _input_values(cfg, op)
        res = core.compute(cfg, op)
        vals.update(res)
        vals["status"] = "Complete" if core.status_of(res) == "complete" else "Incomplete"
        forms = _formulas(r)
        for idx, key in enumerate(KEYS, start=1):
            if key in forms:
                ws_f.cell(r, idx, forms[key])
            else:
                ws_f.cell(r, idx, vals.get(key))
            ws_v.cell(r, idx, vals.get(key))
        for key in core.CATEGORIES:
            for ws in (ws_f, ws_v):
                ws[f"{COL[key]}{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        for key in CALC_KEYS:
            for ws in (ws_f, ws_v):
                ws[f"{COL[key]}{r}"].fill = PatternFill("solid", fgColor=LIGHT)
                ws[f"{COL[key]}{r}"].number_format = "0.00"
        ws_f[f"{COL['area_ha']}{r}"].number_format = "0.##"

    # ---- Check sheet: Operations vs Operations (values)
    last = len(ops) + 1
    ws_c["A1"] = "Overall"
    ws_c["A1"].font = Font(bold=True)
    for idx, key in enumerate(CALC_KEYS, start=3):
        ws_c.cell(3, idx, dict((k, h) for k, h, _ in COLUMNS)[key])
    ws_c.cell(3, 1, "operations")
    ws_c.cell(3, 2, "Week")
    _plain_header(ws_c, 3, 2 + len(CALC_KEYS))
    ws_c.row_dimensions[3].height = 48
    for n, op in enumerate(ops):
        r = n + 2
        cr = n + 4
        ws_c.cell(cr, 1, f"=Operations!{COL['pt']}{r}")
        ws_c.cell(cr, 2, f"=Operations!{COL['week']}{r}")
        for idx, key in enumerate(CALC_KEYS, start=3):
            a, b = f"Operations!{COL[key]}{r}", f"'Operations (values)'!{COL[key]}{r}"
            ws_c.cell(cr, idx, (f'=IF(AND({a}="",{b}=""),"OK",IF(AND(ISNUMBER({a}),ISNUMBER({b})),'
                                f'IF(ABS({a}-{b})<=0.000001,"OK","DIFF"),"DIFF"))'))
    if ops:
        rng = f"C4:{L(2 + len(CALC_KEYS))}{3 + len(ops)}"
        ws_c["B1"] = f'=IF(COUNTIF({rng},"DIFF")=0,"ALL OK","CHECK: "&COUNTIF({rng},"DIFF")&" difference(s)")'
    else:
        ws_c["B1"] = "No operations"
    for idx in range(1, 3 + len(CALC_KEYS)):
        ws_c.column_dimensions[L(idx)].width = 16

    # ---- Normatives
    nrm = norms.get(company) or core.default_norms()
    heads = ["Company", "Work Shift (h)", "Tank per Area (unit)", "Engine Hours (incl. turns) ×",
             "Cost of Diesel (AZN/L)", "Cost of Lubricants (AZN/L)"]
    for idx, h in enumerate(heads, start=1):
        ws_n.cell(1, idx, h)
    _plain_header(ws_n, 1, len(heads))
    ws_n.append([company, nrm["work_shift_h"], nrm["round_up_tanks"], nrm["engine_factor"],
                 nrm["diesel_cost"], nrm["lubricant_cost"]])
    for idx in range(1, 7):
        ws_n.column_dimensions[L(idx)].width = 26

    # ---- inventory
    inv = inventory.get(company) or {"tractors": [], "implements": []}
    for ws, heads, rows in (
        (ws_t, ["Company", "Type of Machinery", "Make / Model", "Quantity", "HP"],
         [[company, r.get("type"), r.get("model"), r.get("quantity"), r.get("hp")] for r in inv["tractors"]]),
        (ws_i, ["Company", "Implement / Equipment", "Brand", "Quantity"],
         [[company, r.get("implement"), r.get("brand"), r.get("quantity")] for r in inv["implements"]]),
    ):
        for idx, h in enumerate(heads, start=1):
            ws.cell(1, idx, h)
        _plain_header(ws, 1, len(heads))
        for row in rows:
            ws.append(row)
        for idx in range(1, len(heads) + 1):
            ws.column_dimensions[L(idx)].width = 26

    # ---- notes
    notes = [
        "Agro Food Investment — export notes",
        f"Company: {company}    Created: {dt.date.today().isoformat()}",
        "",
        "Operations: input values plus live Excel formulas. The formulas follow the logic of the company IPM templates.",
        "Operations (values): the same rows with the values calculated by the application.",
        "Check: compares the two sheets. Cell B1 must read ALL OK after the workbook is opened in Excel.",
        "Every row is self-contained: it carries its own area and normatives, so rows can be copied into a combined workbook.",
        "'Work shift, hours' is the effective shift of the Excel template: work shift minus time for machinery maintenance.",
        "'Cost of diesel / lubricants, azn/liter' are the calculated costs (AZN), as in the Excel templates.",
        "Empty calculated cells mean that an input is missing (the Excel templates show an empty cell as well).",
        "Row length, rows, max duration and speed are informational; no template formula uses them.",
        "Normatives of the selected company are on the 'Normatives' sheet; each operation keeps the values used when it was saved.",
    ]
    for r, line in enumerate(notes, start=1):
        ws_x.cell(r, 1, line)
    ws_x["A1"].font = Font(bold=True, size=13)
    ws_x.column_dimensions["A"].width = 120

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


_TR = str.maketrans({"ə": "e", "Ə": "E", "ğ": "g", "Ğ": "G", "ı": "i", "İ": "I", "ö": "o", "Ö": "O",
                     "ü": "u", "Ü": "U", "ç": "c", "Ç": "C", "ş": "s", "Ş": "S"})


def file_name(company: str, today: Optional[dt.date] = None) -> str:
    slug = company.translate(_TR).replace(" ", "_")
    return f"Agro_Food_Investment_{slug}_{(today or dt.date.today()).isoformat()}.xlsx"
