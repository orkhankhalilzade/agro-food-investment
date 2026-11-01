"""UI-independent logic of the Agro Food Investment planner.

Nothing here imports streamlit, so it can be unit-tested directly.
"""
from __future__ import annotations

import json
import math
import os
import uuid
from dataclasses import asdict
from typing import Any, Optional

from calc import Norms, OperationInputs, calculate, missing_inputs, _is_num, _n

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "config", "orchards.json")

# Exact company -> orchard mapping from the project brief.
COMPANIES = ["Zəfər Bağları", "Grand Agro", "AZROSE", "Leçək", "Şəmkir Agropark"]
ORCHARDS: dict[str, list[Optional[str]]] = {
    "Zəfər Bağları": [None],
    "Grand Agro": ["Almonds", "Olive"],
    "AZROSE": ["Outside Roses", "Peony Outside"],
    "Leçək": ["Damascena"],
    "Şəmkir Agropark": ["Almond", "Cherry", "Persimmon", "Apricot", "Quince", "Pear", "Plum"],
}
STAGES = ["Dormancy", "Bud Swelling", "Flowering", "Fruit Development",
          "Ripening and Harvest", "Harvest", "Harvest and Dormancy"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
CATEGORIES = ["pests", "diseases", "nutrition", "regulators", "adjuvant", "herbicide"]

# Manual inputs stored per operation (informational ones included).
NUMERIC_INPUTS = ["solution_per_ha", "row_length", "rows", "max_duration", "tractors", "speed",
                  "fuel_with_sprayer", "oil_with_sprayer", "fuel_without_sprayer", "oil_without_sprayer",
                  "maintenance_time_h", "tank_capacity_l", "reload_minutes", "productivity_ha_h",
                  "salary_per_hour"]
TEXT_INPUTS = ["tractor"]
NORM_FIELDS = ["work_shift_h", "round_up_tanks", "engine_factor", "diesel_cost", "lubricant_cost"]
SAVE_VERSION = 1


# --------------------------------------------------------------------------- config
class Config:
    def __init__(self, path: str = CONFIG_PATH):
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)
        self._by_key: dict[tuple[str, Optional[str]], dict] = {(o["company"], o["orchard"]): o for o in raw}
        for company, orchards in ORCHARDS.items():
            for orch in orchards:
                if (company, orch) not in self._by_key:
                    raise ValueError(f"Config is missing {company} / {orch}")
        self._calendar: dict[str, dict[str, list[int]]] = {c: self._build_calendar(c) for c in COMPANIES}

    def _entry(self, company, orchard):
        return self._by_key[(company, orchard)]

    def area(self, company: str, orchard: Optional[str]) -> float:
        return float(self._entry(company, orchard)["area_ha"])

    def _build_calendar(self, company: str) -> dict[str, list[int]]:
        # Use the sheet with the most complete month -> week table (Grand Agro's Almonds sheet
        # only lists the first week of each month; its Olive sheet has the full table).
        best = max((self._entry(company, o) for o in ORCHARDS[company]),
                   key=lambda e: sum(len(v) for v in e["months"].values()))
        cal = {m: sorted(best["months"][m]) for m in MONTHS}
        weeks = [w for m in MONTHS for w in cal[m]]
        if weeks != list(range(1, len(weeks) + 1)):
            raise ValueError(f"Calendar of {company} is not a continuous 1..N sequence")
        return cal

    def calendar(self, company: str) -> dict[str, list[int]]:
        return self._calendar[company]

    def weeks_of(self, company: str, month: str) -> list[int]:
        return list(self._calendar[company][month])

    def month_of(self, company: str, week: int) -> Optional[str]:
        for m, ws in self._calendar[company].items():
            if week in ws:
                return m
        return None

    def program(self, company: str, orchard: Optional[str], week: int) -> dict[str, list[str]]:
        return self._entry(company, orchard)["program"].get(str(week), {})

    def has_program(self, company: str, orchard: Optional[str], week: int) -> bool:
        return bool(self.program(company, orchard, week))

    def suggestions(self, company: str, orchard: Optional[str], cat: str) -> list[str]:
        """Distinct program entries of this orchard for one category (text with line breaks)."""
        e = self._entry(company, orchard)
        out: list[str] = []
        for wk in sorted(e["program"], key=lambda x: (not x.isdigit(), int(x) if x.isdigit() else 0)):
            lines = e["program"][wk].get(cat)
            if lines:
                txt = "\n".join(lines)
                if txt not in out:
                    out.append(txt)
        for note in e.get("general_notes", {}).get(cat, []):
            if note not in out:
                out.append(note)
        return out


# --------------------------------------------------------------------------- operations
def default_norms() -> dict[str, float]:
    return asdict(Norms())


def new_id() -> str:
    return uuid.uuid4().hex[:12]


def pt_number(pt: str) -> int:
    return int(pt.split("-")[1])


def next_pt(operations: list[dict], company: str, orchard: Optional[str]) -> str:
    nums = [pt_number(o["pt"]) for o in operations if o["company"] == company and o["orchard"] == orchard]
    return f"PT-{max(nums, default=0) + 1}"


def to_inputs(cfg: Config, op: dict) -> OperationInputs:
    i = op["inputs"]
    return OperationInputs(
        area_ha=cfg.area(op["company"], op["orchard"]),
        solution_per_ha=i.get("solution_per_ha"), fuel_with_sprayer=i.get("fuel_with_sprayer"),
        oil_with_sprayer=i.get("oil_with_sprayer"), fuel_without_sprayer=i.get("fuel_without_sprayer"),
        oil_without_sprayer=i.get("oil_without_sprayer"), maintenance_time_h=i.get("maintenance_time_h"),
        tank_capacity_l=i.get("tank_capacity_l"), productivity_ha_h=i.get("productivity_ha_h"),
        reload_minutes=i.get("reload_minutes"), tractors=i.get("tractors"),
        salary_per_hour=i.get("salary_per_hour"))


def compute(cfg: Config, op: dict) -> dict[str, Optional[float]]:
    return calculate(to_inputs(cfg, op), Norms(**op["norms"]))


def diesel_total(res: dict) -> Optional[float]:
    """Active + transit diesel; None while the active part is missing."""
    if res.get("diesel_active_l") is None:
        return None
    return res["diesel_active_l"] + _n(res.get("diesel_transit_l"))


def status_of(res: dict) -> str:
    return "complete" if res.get("total_cost") is not None else "incomplete"


def sort_key(op: dict):
    return (COMPANIES.index(op["company"]), op["orchard"] or "", op["week"], pt_number(op["pt"]))


def summarize(cfg: Config, operations: list[dict]) -> dict[str, Any]:
    area = hours = diesel = cost = 0.0
    incomplete = 0
    for op in operations:
        res = compute(cfg, op)
        area += cfg.area(op["company"], op["orchard"])
        hours += _n(res["time_h"])
        diesel += _n(diesel_total(res))
        cost += _n(res["total_cost"])
        incomplete += status_of(res) == "incomplete"
    return {"area": area, "count": len(operations), "hours": hours, "diesel": diesel,
            "cost": cost, "incomplete": incomplete}


# --------------------------------------------------------------------------- validation
def _num_or_none(v, name: str, integer: bool = False):
    if v is None:
        return None
    if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or v < 0:
        raise ValueError(f"'{name}' must be a non-negative number or empty")
    if integer and float(v) != int(v):
        raise ValueError(f"'{name}' must be a whole number")
    return float(v)


def validate_operation(cfg: Config, op: dict) -> dict:
    """Return a cleaned copy of ``op`` or raise ValueError."""
    if op.get("company") not in COMPANIES:
        raise ValueError(f"Unknown company: {op.get('company')!r}")
    orchard = op.get("orchard")
    if orchard not in ORCHARDS[op["company"]]:
        raise ValueError(f"Orchard {orchard!r} does not belong to {op['company']}")
    if op.get("stage") not in STAGES:
        raise ValueError(f"Unknown phenological status: {op.get('stage')!r}")
    if op.get("month") not in MONTHS:
        raise ValueError(f"Unknown month: {op.get('month')!r}")
    week = op.get("week")
    if isinstance(week, bool) or not isinstance(week, int) or week not in cfg.weeks_of(op["company"], op["month"]):
        raise ValueError(f"Week {week!r} is not in {op['month']} for {op['company']}")
    pt = op.get("pt")
    if not (isinstance(pt, str) and pt.startswith("PT-") and pt[3:].isdigit() and int(pt[3:]) >= 1):
        raise ValueError(f"Invalid operation number: {pt!r}")
    products = op.get("products") or {}
    inputs_in = op.get("inputs") or {}
    norms_in = op.get("norms") or {}
    cleaned = {
        "id": str(op.get("id") or new_id()), "company": op["company"], "orchard": orchard,
        "stage": op["stage"], "month": op["month"], "week": week, "pt": pt,
        "products": {c: str(products.get(c) or "") for c in CATEGORIES},
        "inputs": {k: _num_or_none(inputs_in.get(k), k, integer=(k in ("rows", "tractors")))
                   for k in NUMERIC_INPUTS},
        "norms": {},
    }
    cleaned["inputs"]["tractor"] = str(inputs_in.get("tractor") or "")
    for k in NORM_FIELDS:
        v = norms_in.get(k, default_norms()[k])
        if k == "round_up_tanks":
            if v not in (0, 1):
                raise ValueError("'round_up_tanks' must be 0 or 1")
            cleaned["norms"][k] = int(v)
        else:
            cleaned["norms"][k] = _num_or_none(v, k)
    return cleaned


# --------------------------------------------------------------------------- save / load
def dump_state(operations: list[dict], norms: dict[str, dict], inventory: dict[str, dict]) -> bytes:
    payload = {"app": "Agro Food Investment", "version": SAVE_VERSION,
               "operations": operations, "norms": norms, "inventory": inventory}
    return json.dumps(payload, ensure_ascii=False, indent=1).encode("utf-8")


def _clean_inventory(inv: Any) -> dict[str, dict]:
    out = {}
    for company in COMPANIES:
        blk = (inv or {}).get(company) or {}
        tr = [{"type": str(r.get("type") or ""), "model": str(r.get("model") or ""),
               "quantity": _num_or_none(r.get("quantity"), "quantity", integer=True),
               "hp": _num_or_none(r.get("hp"), "hp")} for r in blk.get("tractors", [])]
        im = [{"implement": str(r.get("implement") or ""), "brand": str(r.get("brand") or ""),
               "quantity": _num_or_none(r.get("quantity"), "quantity", integer=True)}
              for r in blk.get("implements", [])]
        out[company] = {"tractors": tr, "implements": im}
    return out


def load_state(cfg: Config, data: bytes) -> dict[str, Any]:
    """Parse and validate a saved file. Raises ValueError with a readable message."""
    try:
        payload = json.loads(data.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as e:
        raise ValueError(f"Not a valid saved file: {e}") from e
    if not isinstance(payload, dict) or payload.get("app") != "Agro Food Investment":
        raise ValueError("This file was not created by Agro Food Investment")
    if payload.get("version") != SAVE_VERSION:
        raise ValueError(f"Unsupported file version: {payload.get('version')!r}")
    ops = []
    seen = set()
    for n, op in enumerate(payload.get("operations", []), 1):
        try:
            c = validate_operation(cfg, op)
        except ValueError as e:
            raise ValueError(f"Operation #{n}: {e}") from e
        if c["id"] in seen:
            raise ValueError(f"Operation #{n}: duplicate id")
        seen.add(c["id"])
        ops.append(c)
    norms = {}
    for company in COMPANIES:
        n_in = (payload.get("norms") or {}).get(company) or {}
        base = default_norms()
        for k in NORM_FIELDS:
            v = n_in.get(k, base[k])
            if k == "round_up_tanks":
                if v not in (0, 1):
                    raise ValueError(f"Normatives of {company}: 'round_up_tanks' must be 0 or 1")
                base[k] = int(v)
            else:
                base[k] = _num_or_none(v, k)
        norms[company] = base
    return {"operations": ops, "norms": norms, "inventory": _clean_inventory(payload.get("inventory"))}
