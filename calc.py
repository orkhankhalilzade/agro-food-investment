"""Calculation engine for Agro Food Investment.

Every formula below is a literal translation of the formulas found in the
company Excel templates (operation columns, rows "operations" .. "Total cost").
Cell references in comments are for the Leçək sheet (column C = first operation);
the other sheets use the same layout shifted by a constant row offset (verified
by tests/test_vs_excel.py).

Rules taken from the templates:
* Excel returns "" (empty text) when a required input is missing or invalid.
  Here that is represented by ``None`` -- never by 0.
* ``N(x)`` in Excel turns a missing value into 0; it is used only for the
  optional "transit" parts of cost and for the number of tractors.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field, fields
from typing import Optional

Num = Optional[float]


@dataclass
class Norms:
    """Normatives block (cells B30, B33, B43, B51, B52 of the template)."""
    work_shift_h: Num = 8.0        # B30  Work shift, hours
    round_up_tanks: int = 1        # B33  Tanks per Area, unit (1 = ROUNDUP, 0 = fractional)
    engine_factor: Num = 0.6       # B43  Engine hours (including turns) -- transit factor
    diesel_cost: Num = 1.1         # B51  Cost of diesel, AZN/l
    lubricant_cost: Num = 20.0     # B52  Cost of lubricants, AZN/l


@dataclass
class OperationInputs:
    """Manual inputs of one operation (template rows 19..35, 41, 54)."""
    area_ha: Num = None                   # B18 (orchard area, fixed per orchard)
    solution_per_ha: Num = None           # row 19 Spraying Solution per ha
    fuel_with_sprayer: Num = None         # row 25 l/h
    oil_with_sprayer: Num = None          # row 26 l/h
    fuel_without_sprayer: Num = None      # row 27 l/h (transit/reload)
    oil_without_sprayer: Num = None       # row 28 l/h (transit/reload)
    maintenance_time_h: Num = None        # row 29 time for machinery maintenance
    tank_capacity_l: Num = None           # row 31 Sprayer, tank liter
    productivity_ha_h: Num = None         # row 34 Productivity, ha/hour (manual input in the template)
    reload_minutes: Num = None            # row 35 Time for tank reloading, min
    tractors: Num = None                  # row 41 Tractors (number); blank counts as 1 in working days
    salary_per_hour: Num = None           # row 54 Salary, gross per hour
    # Informational only: these appear in the template but no formula uses them.
    # row_length_m, rows, max_duration_days, tractor name, speed_km_h


RESULT_KEYS = [
    "effective_shift_h", "total_solution_l", "tanks", "reload_total_h", "reload_per_ha",
    "pure_productivity", "output_per_shift_ha", "time_h", "engine_hours", "working_days",
    "diesel_active_l", "lubricant_active_l", "diesel_transit_l", "lubricant_transit_l",
    "diesel_cost", "lubricant_cost", "total_fuel_lubricants", "salary_gross", "total_cost",
]

# Which manual inputs each result needs (derived from the IF(ISNUMBER(..)) guards).
# Used for the "missing input" message. Results that depend on other results are
# resolved transitively in ``missing_inputs``.
DEPENDS = {
    "effective_shift_h": ["maintenance_time_h"],
    "total_solution_l": ["solution_per_ha", "area_ha"],
    "tanks": ["total_solution_l", "tank_capacity_l"],
    "reload_total_h": ["tanks", "reload_minutes"],
    "reload_per_ha": ["reload_total_h", "area_ha"],
    "time_h": ["productivity_ha_h", "reload_total_h", "area_ha"],
    "pure_productivity": ["time_h", "area_ha"],
    "output_per_shift_ha": ["pure_productivity", "effective_shift_h"],
    "engine_hours": ["time_h", "reload_total_h"],
    "working_days": ["time_h", "effective_shift_h"],
    "diesel_active_l": ["time_h", "reload_total_h", "fuel_with_sprayer"],
    "lubricant_active_l": ["time_h", "reload_total_h", "oil_with_sprayer"],
    "diesel_transit_l": ["reload_total_h", "fuel_without_sprayer"],
    "lubricant_transit_l": ["reload_total_h", "oil_without_sprayer"],
    "diesel_cost": ["diesel_active_l"],
    "lubricant_cost": ["lubricant_active_l"],
    "total_fuel_lubricants": ["diesel_cost"],   # or lubricant_cost (either one is enough)
    "salary_gross": ["salary_per_hour", "time_h"],
    "total_cost": ["salary_gross", "total_fuel_lubricants"],
}


def _is_num(x) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool) and not math.isnan(x)


def _n(x) -> float:
    """Excel N(): numbers stay, everything else is 0."""
    return float(x) if _is_num(x) else 0.0


def _roundup0(x: float) -> float:
    """Excel ROUNDUP(x, 0) for non-negative numbers (round away from zero).

    Excel works on 15 significant digits, so tiny binary noise such as
    3.0000000000000004 must not push the result up by one.
    """
    r = round(x, 12)
    return float(math.ceil(r)) if r >= 0 else -float(math.ceil(-r))


def calculate(inp: OperationInputs, norms: Norms) -> dict[str, Num]:
    """Return all results; ``None`` means "missing input" (Excel ``""``)."""
    r: dict[str, Num] = {k: None for k in RESULT_KEYS}
    A = inp.area_ha

    # row 30: =IF(ISNUMBER(C29),$B$30-C29,"")
    if _is_num(inp.maintenance_time_h) and _is_num(norms.work_shift_h):
        r["effective_shift_h"] = norms.work_shift_h - inp.maintenance_time_h

    # row 32: =IF(ISNUMBER(C19),C19*$B$18,"")
    if _is_num(inp.solution_per_ha) and _is_num(A):
        r["total_solution_l"] = inp.solution_per_ha * A

    # row 33: =IF(AND(ISNUMBER(C32),ISNUMBER(C31)),IF(C31>0,IF($B$33=1,ROUNDUP(C32/C31,0),C32/C31),""),"")
    if _is_num(r["total_solution_l"]) and _is_num(inp.tank_capacity_l) and inp.tank_capacity_l > 0:
        q = r["total_solution_l"] / inp.tank_capacity_l
        r["tanks"] = _roundup0(q) if norms.round_up_tanks == 1 else q

    # row 36: =IF(AND(ISNUMBER(C33),ISNUMBER(C35)),C35*C33/60,"")
    if _is_num(r["tanks"]) and _is_num(inp.reload_minutes):
        r["reload_total_h"] = inp.reload_minutes * r["tanks"] / 60

    # row 37: =IF(ISNUMBER(C36),C36/$B$18,"")   (hours per ha)
    if _is_num(r["reload_total_h"]) and _is_num(A) and A != 0:
        r["reload_per_ha"] = r["reload_total_h"] / A

    # row 42: =IF(AND(ISNUMBER(C34),ISNUMBER(C36)),IF(C34>0,$B$18/C34+C36,""),"")
    if (_is_num(inp.productivity_ha_h) and _is_num(r["reload_total_h"])
            and inp.productivity_ha_h > 0 and _is_num(A)):
        r["time_h"] = A / inp.productivity_ha_h + r["reload_total_h"]

    # row 38: =IF(ISNUMBER(C42),IF(C42>0,$B$18/C42,""),"")
    if _is_num(r["time_h"]) and r["time_h"] > 0 and _is_num(A):
        r["pure_productivity"] = A / r["time_h"]

    # row 39: =IF(AND(ISNUMBER(C38),ISNUMBER(C30)),C38*C30,"")
    if _is_num(r["pure_productivity"]) and _is_num(r["effective_shift_h"]):
        r["output_per_shift_ha"] = r["pure_productivity"] * r["effective_shift_h"]

    # row 43: =IF(ISNUMBER(C42),(C42-C36)+(C36*$B$43),"")
    if _is_num(r["time_h"]) and _is_num(norms.engine_factor):
        r["engine_hours"] = (r["time_h"] - r["reload_total_h"]) + r["reload_total_h"] * norms.engine_factor

    # row 44: =IF(AND(ISNUMBER(C42),ISNUMBER(C30)),IF(C30>0,C42/(C30*MAX(1,N(C41))),""),"")
    if _is_num(r["time_h"]) and _is_num(r["effective_shift_h"]) and r["effective_shift_h"] > 0:
        r["working_days"] = r["time_h"] / (r["effective_shift_h"] * max(1.0, _n(inp.tractors)))

    # row 46/47: =IF(AND(ISNUMBER(C42),ISNUMBER(C25)),(C42-C36)*C25,"")
    if _is_num(r["time_h"]) and _is_num(inp.fuel_with_sprayer):
        r["diesel_active_l"] = (r["time_h"] - r["reload_total_h"]) * inp.fuel_with_sprayer
    if _is_num(r["time_h"]) and _is_num(inp.oil_with_sprayer):
        r["lubricant_active_l"] = (r["time_h"] - r["reload_total_h"]) * inp.oil_with_sprayer

    # row 48/49: =IF(AND(ISNUMBER(C36),ISNUMBER(C27)),C36*C27,"")
    if _is_num(r["reload_total_h"]) and _is_num(inp.fuel_without_sprayer):
        r["diesel_transit_l"] = r["reload_total_h"] * inp.fuel_without_sprayer
    if _is_num(r["reload_total_h"]) and _is_num(inp.oil_without_sprayer):
        r["lubricant_transit_l"] = r["reload_total_h"] * inp.oil_without_sprayer

    # row 51: =IF(ISNUMBER(C46),$B$51*(C46+N(C48)),"")
    if _is_num(r["diesel_active_l"]) and _is_num(norms.diesel_cost):
        r["diesel_cost"] = norms.diesel_cost * (r["diesel_active_l"] + _n(r["diesel_transit_l"]))
    # row 52: =IF(ISNUMBER(C47),$B$52*(C47+N(C49)),"")
    if _is_num(r["lubricant_active_l"]) and _is_num(norms.lubricant_cost):
        r["lubricant_cost"] = norms.lubricant_cost * (r["lubricant_active_l"] + _n(r["lubricant_transit_l"]))

    # row 53: =IF(OR(ISNUMBER(C51),ISNUMBER(C52)),N(C51)+N(C52),"")
    if _is_num(r["diesel_cost"]) or _is_num(r["lubricant_cost"]):
        r["total_fuel_lubricants"] = _n(r["diesel_cost"]) + _n(r["lubricant_cost"])

    # row 55: =IF(AND(ISNUMBER(C54),ISNUMBER(C42)),C54*C42,"")   (salary uses Time, not engine hours)
    if _is_num(inp.salary_per_hour) and _is_num(r["time_h"]):
        r["salary_gross"] = inp.salary_per_hour * r["time_h"]

    # row 56: =IF(AND(ISNUMBER(C55),ISNUMBER(C53)),C55+C53,"")
    if _is_num(r["salary_gross"]) and _is_num(r["total_fuel_lubricants"]):
        r["total_cost"] = r["salary_gross"] + r["total_fuel_lubricants"]

    return r


def missing_inputs(result_key: str, inp: OperationInputs, res: dict[str, Num]) -> list[str]:
    """Manual inputs that are still missing for ``result_key`` (empty if it was computed)."""
    if res.get(result_key) is not None:
        return []
    manual = {f.name for f in fields(OperationInputs)}
    out: list[str] = []

    def walk(key: str):
        for dep in DEPENDS.get(key, []):
            if dep in manual:
                if not _is_num(getattr(inp, dep)) and dep not in out:
                    out.append(dep)
            elif res.get(dep) is None:
                walk(dep)

    # "total_fuel_lubricants" needs diesel OR lubricant cost; the diesel branch is reported
    # as the primary path (filling either branch makes the result appear).
    walk(result_key)
    return out
