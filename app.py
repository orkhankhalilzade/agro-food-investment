"""Agro Food Investment — IPM & operational cost planner (Streamlit)."""
from __future__ import annotations

import hashlib
import hmac
import html
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import streamlit as st

import core
import export_xlsx
from calc import Norms, OperationInputs, calculate, missing_inputs
from i18n import LANGS, input_label, month_label, stage_label, tr

st.set_page_config(page_title="Agro Food Investment", page_icon="🌿", layout="wide")

ALL = "__all__"
NONE_ORCHARD = "__none__"

STYLE = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
html, body, [class*="css"], .stApp { font-family: 'IBM Plex Sans', ui-sans-serif, system-ui, sans-serif; }
.block-container { padding-top: 1.6rem; max-width: 1400px; }
div[data-testid="stVerticalBlockBorderWrapper"] { background: #ffffff; border-radius: 8px; }
.blk { display:flex; gap:.75rem; align-items:center; margin:.6rem 0 .4rem 0; }
.blk .num { width:1.9rem; height:1.9rem; border-radius:50%; background:#144927; color:#f5faf4;
            display:flex; align-items:center; justify-content:center; font-weight:600; font-size:.95rem; flex:none; }
.blk .t { font-size:1.2rem; font-weight:600; color:#121e14; line-height:1.2; }
.blk .s { font-size:.85rem; color:#5c675d; }
.sub { font-weight:600; color:#12361e; margin:.1rem 0 .4rem 0; }
.cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:.75rem; margin:.3rem 0 .6rem 0; }
.card { border:1px solid #dee3de; border-radius:8px; background:#fff; padding:.7rem .9rem; }
.card.primary { background:#144927; border-color:#144927; }
.card .l { font-size:.78rem; color:#5c675d; }
.card .v { font-family:'IBM Plex Mono',ui-monospace,monospace; font-variant-numeric:tabular-nums;
           font-size:1.35rem; font-weight:500; color:#121e14; }
.card .v small { font-size:.8rem; color:#5c675d; }
.card.primary .l, .card.primary .v, .card.primary .v small { color:#f5faf4; }
.card .v.miss, .miss { color:#6e4108; font-size:.9rem; font-family:'IBM Plex Sans',sans-serif; }
.card.primary .v.miss { color:#fee5b3; }
table.res { width:100%; border-collapse:collapse; }
table.res td { padding:.28rem 0; border-bottom:1px solid #ebf2eb; font-size:.9rem; vertical-align:baseline; }
table.res td.k { color:#5c675d; padding-right:.6rem; }
table.res td.v { text-align:right; font-family:'IBM Plex Mono',ui-monospace,monospace; font-variant-numeric:tabular-nums; white-space:nowrap; }
.note { font-size:.8rem; color:#5c675d; }
.warnbox { background:#fee5b3; color:#6e4108; border-radius:6px; padding:.35rem .6rem; font-size:.82rem; margin-top:.4rem; }
</style>
"""


# --------------------------------------------------------------------------- state
@st.cache_resource
def get_cfg() -> core.Config:
    return core.Config()


CFG = get_cfg()
ss = st.session_state


def norm_key(field: str) -> str:
    """One widget key per normative; the values of every company live in ``ss.norms``.

    (Streamlit forgets widgets that are not rendered, so per-company widget keys would lose
    their values whenever another company is selected.)
    """
    return f"n_{field}"


def _norm_value(field: str, v):
    return int(v) if field == "round_up_tanks" else float(v)


def sync_norms():
    """Keep ``ss.norms`` and the normative widgets consistent; runs at the start of every run."""
    cur = ss.norm_company
    ss.norms[cur] = {f: _norm_value(f, ss[norm_key(f)]) for f in core.NORM_FIELDS}
    want = ss.f_company
    if want != cur:
        for f, v in ss.norms[want].items():
            ss[norm_key(f)] = _norm_value(f, v)
        ss.norm_company = want


def init_state():
    ss.setdefault("lang", "EN")
    ss.setdefault("operations", [])
    ss.setdefault("edit_id", None)
    ss.setdefault("confirm_delete", None)
    ss.setdefault("flash", None)
    ss.setdefault("loaded_sig", None)
    ss.setdefault("inv_ver", 0)
    ss.setdefault("inv_init", {})
    ss.setdefault("inv_out", {})
    for k in ("f_company", "f_orchard", "f_stage", "f_month", "f_week", "i_tractor"):
        ss.setdefault(k, "")
    ss.setdefault("i_tractors", 1)
    ss.setdefault("norms", {c: core.default_norms() for c in core.COMPANIES + [""]})
    ss.setdefault("norm_company", "")
    for f, v in core.default_norms().items():
        ss.setdefault(norm_key(f), _norm_value(f, v))
    for k in ("wf_company", "wf_orchard", "wf_month", "wf_week", "wf_op"):
        ss.setdefault(k, ALL)


def T(key: str, **kw) -> str:
    return tr(ss.get("lang") or "EN", key, **kw)


def lang() -> str:
    return ss.get("lang") or "EN"


def esc(s) -> str:
    return html.escape(str(s))


def orch_value(v: str):
    return None if v == NONE_ORCHARD else v


def orch_label(v) -> str:
    return T("none_orchard") if v in (None, NONE_ORCHARD) else str(v)


# --------------------------------------------------------------------------- callbacks
def cb_lang():
    if ss.lang is None:
        ss.lang = "EN"


def cb_company():
    opts = core.ORCHARDS.get(ss.f_company, [])
    ss.f_orchard = (opts[0] or NONE_ORCHARD) if len(opts) == 1 else ""
    ss.f_week = ""


def cb_month():
    ss.f_week = ""


def cb_reset():
    for name in core.NUMERIC_INPUTS:
        ss.pop(f"i_{name}", None)
    ss["i_tractors"] = 1
    ss["i_tractor"] = ""
    for k in [k for k in ss.keys() if k.startswith(("p_", "s_"))]:
        del ss[k]
    ss.edit_id = None


def cb_suggest(pkey: str, skey: str):
    if ss.get(skey):
        ss[pkey] = ss[skey]
    ss[skey] = ""


def current_norms(company: str | None = None) -> dict:
    """Normatives shown in block 2 (they belong to the selected company)."""
    return {f: _norm_value(f, ss[norm_key(f)]) for f in core.NORM_FIELDS}


def form_op_dict(pt: str) -> dict:
    company, orchard = ss.f_company, orch_value(ss.f_orchard)
    inputs = {n: ss.get(f"i_{n}") for n in core.NUMERIC_INPUTS}
    inputs["tractor"] = (ss.get("i_tractor") or "").strip()
    okey = ss.f_orchard if ss.f_orchard else "x"
    products = {c: ss.get(f"p_{c}_{okey}_{ss.f_week}", "") for c in core.CATEGORIES}
    return {"company": company, "orchard": orchard, "stage": ss.f_stage, "month": ss.f_month,
            "week": ss.f_week if ss.f_week != "" else None, "pt": pt,
            "products": products, "inputs": inputs, "norms": current_norms()}


def cb_save():
    required = [ss.f_company, ss.f_orchard, ss.f_stage, ss.f_month, ss.f_week]
    if any(v == "" for v in required):
        ss.flash = ("warning", T("need_fields"))
        return
    editing = next((o for o in ss.operations if o["id"] == ss.edit_id), None)
    company, orchard = ss.f_company, orch_value(ss.f_orchard)
    pt = editing["pt"] if editing else core.next_pt(ss.operations, company, orchard)
    try:
        op = core.validate_operation(CFG, {**form_op_dict(pt), "id": editing["id"] if editing else core.new_id()})
    except ValueError as e:
        ss.flash = ("error", str(e))
        return
    if editing:
        ss.operations = [op if o["id"] == editing["id"] else o for o in ss.operations]
        ss.flash = ("success", T("updated", pt=pt))
        ss.edit_id = None
    else:
        ss.operations = ss.operations + [op]
        ss.flash = ("success", T("added", pt=pt))


def cb_edit(op_id: str):
    op = next((o for o in ss.operations if o["id"] == op_id), None)
    if not op:
        return
    ss.norms[ss.norm_company] = current_norms()          # keep what is on screen for the old company
    ss.f_company = op["company"]
    ss.f_orchard = op["orchard"] or NONE_ORCHARD
    ss.f_stage, ss.f_month, ss.f_week = op["stage"], op["month"], op["week"]
    ss.norms[op["company"]] = dict(op["norms"])
    for f, v in op["norms"].items():
        ss[norm_key(f)] = _norm_value(f, v)
    ss.norm_company = op["company"]
    for name in core.NUMERIC_INPUTS:
        v = op["inputs"].get(name)
        if v is None:
            ss.pop(f"i_{name}", None)
        else:
            ss[f"i_{name}"] = int(v) if name in ("rows", "tractors") else float(v)
    ss["i_tractor"] = op["inputs"].get("tractor", "")
    for k in [k for k in ss.keys() if k.startswith(("p_", "s_"))]:
        del ss[k]
    for c, txt in op["products"].items():
        ss[f"p_{c}_{ss.f_orchard}_{op['week']}"] = txt
    ss.edit_id = op_id
    ss.flash = ("info", T("editing", pt=op["pt"]))


def cb_delete_ask(op_id: str):
    ss.confirm_delete = op_id


def cb_delete_do():
    op = next((o for o in ss.operations if o["id"] == ss.confirm_delete), None)
    if op:
        ss.operations = [o for o in ss.operations if o["id"] != op["id"]]
        if ss.edit_id == op["id"]:
            ss.edit_id = None
        ss.flash = ("success", T("deleted", pt=op["pt"]))
    ss.confirm_delete = None


def cb_delete_cancel():
    ss.confirm_delete = None


# --------------------------------------------------------------------------- inventory helpers
def empty_df(kind: str) -> pd.DataFrame:
    if kind == "tractors":
        return pd.DataFrame({"type": pd.Series(dtype="object"), "model": pd.Series(dtype="object"),
                             "quantity": pd.Series(dtype="float64"), "hp": pd.Series(dtype="float64")})
    return pd.DataFrame({"implement": pd.Series(dtype="object"), "brand": pd.Series(dtype="object"),
                         "quantity": pd.Series(dtype="float64")})


def df_records(df: pd.DataFrame, kind: str) -> list[dict]:
    out = []
    for _, r in df.iterrows():
        rec = {}
        for col in df.columns:
            v = r[col]
            rec[col] = None if (v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NA) else v
        texts = [str(rec.get(c) or "").strip() for c in (("type", "model") if kind == "tractors" else ("implement", "brand"))]
        if not any(texts) and rec.get("quantity") is None and rec.get("hp") is None:
            continue
        rec["quantity"] = None if rec.get("quantity") is None else int(rec["quantity"])
        out.append(rec)
    return out


def inventory_snapshot() -> dict:
    inv = {}
    for c in core.COMPANIES:
        inv[c] = {"tractors": [], "implements": []}
        for kind in ("tractors", "implements"):
            df = ss.inv_out.get((c, kind))
            if df is not None:
                inv[c][kind] = df_records(df, kind)
    return inv


def norms_snapshot() -> dict:
    return {c: dict(ss.norms[c]) for c in core.COMPANIES}


# --------------------------------------------------------------------------- save / load
def handle_load(uploaded):
    data = uploaded.getvalue()
    sig = hashlib.sha1(data).hexdigest()
    if sig == ss.loaded_sig:
        return
    ss.loaded_sig = sig
    try:
        state = core.load_state(CFG, data)
    except ValueError as e:
        ss.flash = ("error", T("load_bad", err=str(e)))
        return
    ss.operations = state["operations"]
    for c, n in state["norms"].items():
        ss.norms[c] = dict(n)
    ss.norms[""] = core.default_norms()
    for f, v in ss.norms[ss.f_company].items():
        ss[norm_key(f)] = _norm_value(f, v)
    ss.norm_company = ss.f_company
    ss.inv_ver += 1
    ss.inv_init = {}
    for c, blk in state["inventory"].items():
        for kind in ("tractors", "implements"):
            df = pd.DataFrame(blk[kind]) if blk[kind] else empty_df(kind)
            for col in empty_df(kind).columns:
                if col not in df.columns:
                    df[col] = None
            df = df[list(empty_df(kind).columns)]
            for col in ("quantity", "hp"):
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").astype("float64")
            ss.inv_init[(c, kind)] = df
    ss.edit_id = None
    ss.confirm_delete = None
    ss.flash = ("success", T("load_ok", n=len(state["operations"])))


# --------------------------------------------------------------------------- rendering helpers
def block_head(n: int, title: str, sub: str):
    st.markdown(f'<div class="blk"><span class="num">{n}</span><div><div class="t">{esc(title)}</div>'
                f'<div class="s">{esc(sub)}</div></div></div>', unsafe_allow_html=True)


def sub_head(text: str):
    st.markdown(f'<div class="sub">{esc(text)}</div>', unsafe_allow_html=True)


def num_input(label: str, name: str, *, step: float = 0.1, integer: bool = False, help: str | None = None,
              disabled: bool = False):
    key = f"i_{name}"
    kw = dict(key=key, step=1 if integer else step, format="%d" if integer else "%g", help=help, disabled=disabled)
    mn = 0 if integer else 0.0
    if key in ss:
        return st.number_input(label, min_value=mn, **kw)
    return st.number_input(label, min_value=mn, value=None, **kw)


UNITS = {
    "L": ("L", "L"), "h": ("h", "saat"), "ha": ("ha", "ha"), "ha/h": ("ha/h", "ha/saat"), "h/ha": ("h/ha", "saat/ha"),
    "d": ("d", "gün"), "pcs": ("pcs", "ədəd"), "AZN": ("AZN", "AZN"), "": ("", ""),
}


def unit(u: str) -> str:
    return UNITS[u][1 if lang() == "AZ" else 0]


def fmt(v: float, dec: int = 2) -> str:
    return f"{v:,.{dec}f}"


def value_html(v, u: str, dec: int = 2) -> str:
    if v is None:
        return f'<span class="miss">{esc(T("missing"))}</span>'
    return f'{fmt(v, dec)} <small>{esc(unit(u))}</small>'


def card(label: str, v, u: str, primary: bool = False, dec: int = 2) -> str:
    cls = "card primary" if primary else "card"
    vcls = "v miss" if v is None else "v"
    inner = esc(T("missing")) if v is None else value_html(v, u, dec)
    return f'<div class="{cls}"><div class="l">{esc(label)}</div><div class="{vcls}">{inner}</div></div>'


def missing_caption(keys: list[str], inp: OperationInputs, res: dict) -> str:
    need: list[str] = []
    for k in keys:
        if res.get(k) is None:
            for m in missing_inputs(k, inp, res):
                lab = input_label(lang(), m)
                if lab not in need:
                    need.append(lab)
    if not need:
        return ""
    return f'<div class="warnbox">{esc(T("missing_list", list=", ".join(need)))}</div>'


def result_group(title_key: str, rows: list[tuple[str, str, str, int]], inp: OperationInputs, res: dict):
    with st.container(border=True):
        sub_head(T(title_key))
        body = "".join(f'<tr><td class="k">{esc(T(lab))}</td><td class="v">{value_html(res[k], u, d)}</td></tr>'
                       for lab, k, u, d in rows)
        st.markdown(f'<table class="res">{body}</table>', unsafe_allow_html=True)
        cap = missing_caption([k for _, k, _, _ in rows], inp, res)
        if cap:
            st.markdown(cap, unsafe_allow_html=True)


# --------------------------------------------------------------------------- page
def render_header():
    left, mid, right = st.columns([5, 4, 2], vertical_alignment="center")
    # The uploaded file must be processed BEFORE anything that reads the data (save button,
    # widgets below), so it is handled first; the layout slots above only fix where things appear.
    up_box = st.container()
    with up_box:
        up = st.file_uploader(T("load_lbl"), type=["json"], key="uploader")
    if up is not None:
        handle_load(up)
    else:
        ss.loaded_sig = None
    with left:
        st.markdown(f'<div class="t" style="font-size:1.6rem;font-weight:600;color:#144927">{esc(T("app_title"))}</div>'
                    f'<div class="s" style="color:#5c675d">{esc(T("app_subtitle"))}</div>', unsafe_allow_html=True)
    with mid:
        st.download_button(T("save_btn"), data=core.dump_state(ss.operations, norms_snapshot(), inventory_snapshot()),
                           file_name="agro_food_investment_progress.json", mime="application/json",
                           help=T("save_note"), key="save_btn")
    with right:
        st.segmented_control(T("language"), LANGS, key="lang", on_change=cb_lang, label_visibility="collapsed")


def render_block1():
    block_head(1, T("b1"), T("b1_sub"))
    with st.container(border=True):
        c = st.columns(6)
        with c[0]:
            st.selectbox(T("company"), [""] + core.COMPANIES, key="f_company", on_change=cb_company,
                         format_func=lambda v: T("select") if v == "" else v)
        company = ss.f_company
        orch_opts = [""] + [(o or NONE_ORCHARD) for o in core.ORCHARDS.get(company, [])]
        if ss.f_orchard not in orch_opts:
            ss.f_orchard = ""
        if len(orch_opts) == 2 and ss.f_orchard == "":
            ss.f_orchard = orch_opts[1]
        with c[1]:
            st.selectbox(T("orchard"), orch_opts, key="f_orchard",
                         format_func=lambda v: T("select") if v == "" else orch_label(v))
        with c[2]:
            area_txt = ""
            if company and ss.f_orchard:
                area_txt = f"{CFG.area(company, orch_value(ss.f_orchard)):g}"
            st.text_input(T("area"), value=area_txt, disabled=True, help=T("area_help"))
        with c[3]:
            st.selectbox(T("stage"), [""] + core.STAGES, key="f_stage",
                         format_func=lambda v: T("select") if v == "" else stage_label(lang(), v))
        with c[4]:
            st.selectbox(T("month"), [""] + core.MONTHS, key="f_month", on_change=cb_month,
                         format_func=lambda v: T("select") if v == "" else month_label(lang(), v))
        week_opts = [""] + (CFG.weeks_of(company, ss.f_month) if company and ss.f_month else [])
        if ss.f_week not in week_opts:
            ss.f_week = ""
        with c[5]:
            st.selectbox(T("week"), week_opts, key="f_week",
                         format_func=lambda v: T("select") if v == "" else f"{T('week')} {v}")


def render_block2():
    block_head(2, T("b2"), T("b2_sub"))
    with st.container(border=True):
        c = st.columns(5)
        with c[0]:
            st.number_input(T("n_shift"), min_value=0.0, step=0.5, format="%g", key=norm_key("work_shift_h"))
        with c[1]:
            st.selectbox(T("n_tank"), [1, 0], key=norm_key("round_up_tanks"), help=T("n_tank_help"))
        with c[2]:
            st.number_input(T("n_engine"), min_value=0.0, step=0.05, format="%g", key=norm_key("engine_factor"))
        with c[3]:
            st.number_input(T("n_diesel"), min_value=0.0, step=0.05, format="%g", key=norm_key("diesel_cost"))
        with c[4]:
            st.number_input(T("n_lub"), min_value=0.0, step=0.5, format="%g", key=norm_key("lubricant_cost"))


def render_products():
    company, orchard, week = ss.f_company, ss.f_orchard, ss.f_week
    ready = bool(company and orchard and week != "")
    okey = orchard if orchard else "x"
    if not ready:
        st.markdown(f'<div class="note">{esc(T("program_pick"))}</div>', unsafe_allow_html=True)
    else:
        has = CFG.has_program(company, orch_value(orchard), week)
        st.markdown(f'<div class="note">{esc(T("program_loaded") if has else T("program_none"))}</div>',
                    unsafe_allow_html=True)
    for cat in core.CATEGORIES:
        pkey, skey = f"p_{cat}_{okey}_{week}", f"s_{cat}_{okey}_{week}"
        if ready and pkey not in ss:
            ss[pkey] = "\n".join(CFG.program(company, orch_value(orchard), week).get(cat, []))
        elif not ready:
            ss.setdefault(pkey, "")
        st.text_area(T(cat), key=pkey, height=96, disabled=not ready)
        if ready:
            sugg = CFG.suggestions(company, orch_value(orchard), cat)
            if sugg:
                ss.setdefault(skey, "")
                st.selectbox(T("suggest"), [""] + sugg, key=skey, label_visibility="collapsed",
                             format_func=lambda v: T("suggest") + " …" if v == "" else " | ".join(v.split("\n"))[:90],
                             on_change=cb_suggest, args=(pkey, skey))


def render_block3():
    block_head(3, T("b3"), T("b3_sub"))
    company, orchard = ss.f_company, orch_value(ss.f_orchard)
    editing = next((o for o in ss.operations if o["id"] == ss.edit_id), None)
    if editing:
        st.info(T("editing", pt=editing["pt"]))
    r1 = st.columns(3)
    with r1[0], st.container(border=True):
        sub_head(T("op_info"))
        if editing:
            pt = editing["pt"]
        elif company and ss.f_orchard:
            pt = core.next_pt(ss.operations, company, orchard)
        else:
            pt = "—"
        st.text_input(T("operation"), value=T("operation_auto", pt=pt) if pt != "—" else "—", disabled=True)
        num_input(T("solution"), "solution_per_ha", step=10.0)
        num_input(T("row_len"), "row_length", step=1.0)
        num_input(T("rows"), "rows", integer=True)
        num_input(T("max_dur"), "max_duration", step=1.0)
        st.markdown(f'<div class="note">{esc(T("info_note"))}</div>', unsafe_allow_html=True)
    with r1[1], st.container(border=True):
        sub_head(T("mach"))
        st.text_input(T("tractor"), key="i_tractor", placeholder=T("tractor_ph"))
        num_input(T("n_tractors"), "tractors", integer=True)
        num_input(T("speed"), "speed", step=0.5)
        num_input(T("fuel_w"), "fuel_with_sprayer", step=0.5)
        num_input(T("oil_w"), "oil_with_sprayer", step=0.01)
        num_input(T("fuel_wo"), "fuel_without_sprayer", step=0.5)
        num_input(T("oil_wo"), "oil_without_sprayer", step=0.01)
        num_input(T("maint"), "maintenance_time_h", step=0.25)
    with r1[2], st.container(border=True):
        sub_head(T("ipm"))
        render_products()
    r2 = st.columns(3)
    with r2[0], st.container(border=True):
        sub_head(T("sprayer"))
        num_input(T("tank_cap"), "tank_capacity_l", step=50.0)
        num_input(T("reload"), "reload_minutes", step=1.0)
        num_input(T("prod"), "productivity_ha_h", step=0.1, help=T("prod_help"))
    with r2[1], st.container(border=True):
        sub_head(T("labor"))
        num_input(T("salary"), "salary_per_hour", step=0.1)
    with r2[2]:
        b1, b2 = st.columns(2)
        b1.button(T("reset"), on_click=cb_reset, width="stretch")
        b2.button(T("update") if editing else T("add"), on_click=cb_save, type="primary", width="stretch")
        if editing:
            st.button(T("cancel_edit"), on_click=cb_reset, width="stretch")
        if ss.flash:
            kind, msg = ss.flash
            getattr(st, kind)(msg)
            ss.flash = None
        if company:
            n_co = sum(1 for o in ss.operations if o["company"] == company)
            if n_co:
                st.markdown(f'<div class="note">{esc(T("count_info", n=n_co, company=company))}</div>',
                            unsafe_allow_html=True)


def render_block4():
    block_head(4, T("b4"), T("b4_sub"))
    company, orchard = ss.f_company, orch_value(ss.f_orchard)
    area = CFG.area(company, orchard) if company and ss.f_orchard else None
    g = lambda n: ss.get(f"i_{n}")
    inp = OperationInputs(
        area_ha=area, solution_per_ha=g("solution_per_ha"), fuel_with_sprayer=g("fuel_with_sprayer"),
        oil_with_sprayer=g("oil_with_sprayer"), fuel_without_sprayer=g("fuel_without_sprayer"),
        oil_without_sprayer=g("oil_without_sprayer"), maintenance_time_h=g("maintenance_time_h"),
        tank_capacity_l=g("tank_capacity_l"), reload_minutes=g("reload_minutes"),
        productivity_ha_h=g("productivity_ha_h"), tractors=g("tractors"), salary_per_hour=g("salary_per_hour"))
    res = calculate(inp, Norms(**current_norms()))
    cards = "".join([
        card(T("r_total_cost"), res["total_cost"], "AZN", primary=True),
        card(T("r_salary"), res["salary_gross"], "AZN"),
        card(T("r_fuel_total"), res["total_fuel_lubricants"], "AZN"),
        card(T("r_working_days"), res["working_days"], "d"),
    ])
    st.markdown(f'<div class="cards">{cards}</div>', unsafe_allow_html=True)
    cap = missing_caption(["total_cost", "working_days"], inp, res)
    if cap:
        st.markdown(cap, unsafe_allow_html=True)
    c = st.columns(3)
    with c[0]:
        result_group("g_solution", [("r_total_solution", "total_solution_l", "L", 2), ("r_tanks", "tanks", "pcs", 2),
                                    ("r_reload_total", "reload_total_h", "h", 2),
                                    ("r_reload_per_ha", "reload_per_ha", "h/ha", 3)], inp, res)
    with c[1]:
        result_group("g_prod", [("r_pure", "pure_productivity", "ha/h", 2), ("r_output", "output_per_shift_ha", "ha", 2),
                                ("r_time", "time_h", "h", 2), ("r_engine", "engine_hours", "h", 2),
                                ("r_working_days", "working_days", "d", 2)], inp, res)
    with c[2]:
        result_group("g_fuel", [("r_diesel_active", "diesel_active_l", "L", 2), ("r_lub_active", "lubricant_active_l", "L", 2),
                                ("r_diesel_transit", "diesel_transit_l", "L", 2), ("r_lub_transit", "lubricant_transit_l", "L", 2),
                                ("r_diesel_cost", "diesel_cost", "AZN", 2), ("r_lub_cost", "lubricant_cost", "AZN", 2),
                                ("r_fuel_total", "total_fuel_lubricants", "AZN", 2)], inp, res)


def filtered_operations() -> list[dict]:
    ops = sorted(ss.operations, key=core.sort_key)
    f = ss
    out = []
    for o in ops:
        if f.wf_company != ALL and o["company"] != f.wf_company:
            continue
        if f.wf_orchard != ALL and (o["orchard"] or NONE_ORCHARD) != f.wf_orchard:
            continue
        if f.wf_month != ALL and o["month"] != f.wf_month:
            continue
        if f.wf_week != ALL and o["week"] != f.wf_week:
            continue
        if f.wf_op != ALL and o["pt"] != f.wf_op:
            continue
        out.append(o)
    return out


def render_block5():
    block_head(5, T("b5"), T("b5_sub"))
    ops_all = ss.operations
    with st.container(border=True):
        c = st.columns([2, 2, 2, 2, 2, 2], vertical_alignment="bottom")
        company_opts = [ALL] + core.COMPANIES
        if ss.wf_company not in company_opts:
            ss.wf_company = ALL
        with c[0]:
            st.selectbox(T("company"), company_opts, key="wf_company",
                         format_func=lambda v: T("all") if v == ALL else v)
        orch_all = [(o or NONE_ORCHARD) for co in core.COMPANIES if ss.wf_company in (ALL, co) for o in core.ORCHARDS[co]]
        orch_opts = [ALL] + list(dict.fromkeys(orch_all))
        if ss.wf_orchard not in orch_opts:
            ss.wf_orchard = ALL
        with c[1]:
            st.selectbox(T("orchard"), orch_opts, key="wf_orchard",
                         format_func=lambda v: T("all") if v == ALL else orch_label(v))
        if ss.wf_month not in [ALL] + core.MONTHS:
            ss.wf_month = ALL
        with c[2]:
            st.selectbox(T("month"), [ALL] + core.MONTHS, key="wf_month",
                         format_func=lambda v: T("all") if v == ALL else month_label(lang(), v))
        week_opts = [ALL] + sorted({o["week"] for o in ops_all})
        if ss.wf_week not in week_opts:
            ss.wf_week = ALL
        with c[3]:
            st.selectbox(T("week"), week_opts, key="wf_week",
                         format_func=lambda v: T("all") if v == ALL else f"{T('week')} {v}")
        pt_opts = [ALL] + sorted({o["pt"] for o in ops_all}, key=core.pt_number)
        if ss.wf_op not in pt_opts:
            ss.wf_op = ALL
        with c[4]:
            st.selectbox(T("operation"), pt_opts, key="wf_op", format_func=lambda v: T("all") if v == ALL else v)
        with c[5]:
            companies_with_ops = sorted({o["company"] for o in ops_all}, key=core.COMPANIES.index)
            if ss.wf_company != ALL:
                exp_company = ss.wf_company
            elif len(companies_with_ops) == 1:
                exp_company = companies_with_ops[0]
            elif ss.f_company in companies_with_ops:
                exp_company = ss.f_company
            else:
                exp_company = ""
            n_exp = sum(1 for o in ops_all if o["company"] == exp_company)
            if exp_company and n_exp:
                data = export_xlsx.build_workbook(CFG, exp_company, ops_all, norms_snapshot(), inventory_snapshot())
                st.download_button(T("export"), data=data, file_name=export_xlsx.file_name(exp_company),
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="export_btn", type="primary", width="stretch")
            else:
                st.button(T("export"), disabled=True, width="stretch", key="export_dis",
                          help=T("export_empty") if not companies_with_ops else T("export_none"))
        if exp_company and n_exp:
            st.markdown(f'<div class="note">{esc(T("export_info", n=n_exp, company=exp_company))}</div>',
                        unsafe_allow_html=True)
        elif companies_with_ops:
            st.markdown(f'<div class="note">{esc(T("export_none"))}</div>', unsafe_allow_html=True)

    ops = filtered_operations()
    s = core.summarize(CFG, ops)
    cards = "".join([
        card(T("st_area"), s["area"], "ha"), card(T("st_ops"), float(s["count"]), "", dec=0),
        card(T("st_hours"), s["hours"], "h", dec=1), card(T("st_diesel"), s["diesel"], "L", dec=1),
        card(T("st_cost"), s["cost"], "AZN"),
    ]).replace("<small></small>", "")
    st.markdown(f'<div class="cards">{cards}</div>', unsafe_allow_html=True)
    if s["incomplete"]:
        st.markdown(f'<div class="note">{esc(T("st_incomplete", n=s["incomplete"]))}</div>', unsafe_allow_html=True)

    if not ops:
        st.info(T("empty_table"))
        return
    rows = []
    for o in ops:
        res = core.compute(CFG, o)
        rows.append({
            T("col_company"): o["company"], T("col_orchard"): orch_label(o["orchard"]), T("col_week"): o["week"],
            T("col_operation"): o["pt"], T("col_area"): CFG.area(o["company"], o["orchard"]),
            T("col_tractor"): o["inputs"].get("tractor") or "—", T("col_days"): res["working_days"],
            T("col_diesel"): core.diesel_total(res), T("col_cost"): res["total_cost"],
            T("col_status"): T("status_" + core.status_of(res))})
    df = pd.DataFrame(rows)
    numcfg = {c: st.column_config.NumberColumn(format="%.2f") for c in
              (T("col_area"), T("col_days"), T("col_diesel"), T("col_cost"))}
    ev = st.dataframe(df, hide_index=True, width="stretch", on_select="rerun", selection_mode="single-row",
                      column_config=numcfg, key="wk_table")
    sel = ev.selection.rows if ev is not None else []
    if sel and sel[0] < len(ops):
        op = ops[sel[0]]
        b = st.columns([1, 1, 6])
        b[0].button(T("edit"), on_click=cb_edit, args=(op["id"],), key="btn_edit", width="stretch")
        b[1].button(T("delete"), on_click=cb_delete_ask, args=(op["id"],), key="btn_del", width="stretch")
    else:
        st.markdown(f'<div class="note">{esc(T("select_row"))}</div>', unsafe_allow_html=True)
    cd = next((o for o in ss.operations if o["id"] == ss.confirm_delete), None)
    if cd:
        st.warning(T("confirm_delete", pt=cd["pt"], orchard=orch_label(cd["orchard"]), week=cd["week"]))
        k = st.columns([1, 1, 6])
        k[0].button(T("confirm"), on_click=cb_delete_do, key="btn_del_yes", type="primary", width="stretch")
        k[1].button(T("cancel"), on_click=cb_delete_cancel, key="btn_del_no", width="stretch")


def render_inventory():
    st.markdown(f'<div class="note">{esc(T("inv_help"))}</div>', unsafe_allow_html=True)
    tabs = st.tabs(core.COMPANIES)
    cfgs = {
        "tractors": {
            "type": st.column_config.TextColumn("Type of Machinery / Texnikanın növü"),
            "model": st.column_config.TextColumn("Make / Model / Marka, model"),
            "quantity": st.column_config.NumberColumn("Quantity / Sayı", min_value=0, step=1, format="%d"),
            "hp": st.column_config.NumberColumn("HP / At gücü", min_value=0, format="%g")},
        "implements": {
            "implement": st.column_config.TextColumn("Implement / Equipment / Qoşqu, avadanlıq"),
            "brand": st.column_config.TextColumn("Brand / Marka"),
            "quantity": st.column_config.NumberColumn("Quantity / Sayı", min_value=0, step=1, format="%d")},
    }
    for tab, company in zip(tabs, core.COMPANIES):
        with tab:
            for kind, title in (("tractors", "inv_tractors"), ("implements", "inv_impl")):
                sub_head(T(title))
                init = ss.inv_init.get((company, kind))
                if init is None:
                    init = ss.inv_init[(company, kind)] = empty_df(kind)
                edited = st.data_editor(init, num_rows="dynamic", hide_index=True, width="stretch",
                                        column_config=cfgs[kind], key=f"inv_{kind}_{company}_{ss.inv_ver}")
                ss.inv_out[(company, kind)] = edited


# --------------------------------------------------------------------------- access control
def _configured_password() -> str:
    pw = os.environ.get("AGRO_PASSWORD", "")
    if not pw:
        try:
            pw = str(st.secrets.get("password", ""))
        except Exception:        # no secrets file
            pw = ""
    return pw


def require_password() -> bool:
    """Shared-password gate. Secure by default: without a configured password the app stays closed
    (set AGRO_ALLOW_OPEN=1 only for local use on your own computer)."""
    expected = _configured_password()
    if not expected:
        if os.environ.get("AGRO_ALLOW_OPEN") == "1":
            return True
        st.error("Password is not configured. Set the AGRO_PASSWORD environment variable. / "
                 "Şifrə təyin edilməyib. AGRO_PASSWORD mühit dəyişənini təyin edin.")
        return False
    if ss.get("authed"):
        return True
    st.markdown('<div class="t" style="font-size:1.5rem;font-weight:600;color:#144927">Agro Food Investment</div>',
                unsafe_allow_html=True)
    with st.form("login"):
        pw = st.text_input("Password / Şifrə", type="password")
        go = st.form_submit_button("Enter / Daxil ol", type="primary")
    if go:
        if hmac.compare_digest(pw.encode("utf-8"), expected.encode("utf-8")):
            ss["authed"] = True
            ss["login_fails"] = 0
            st.rerun()
        ss["login_fails"] = ss.get("login_fails", 0) + 1
        time.sleep(min(2 ** ss["login_fails"], 30) / 2)       # slow down guessing
        st.error("Wrong password. / Şifrə yanlışdır.")
    return False


def main():
    st.markdown(STYLE, unsafe_allow_html=True)
    if not require_password():
        return
    init_state()
    sync_norms()
    render_header()
    # static bilingual labels: the selected tab must not jump when the language changes
    tab_plan, tab_inv = st.tabs(["Planning · Planlaşdırma", "Machinery & Equipment · Texnika və avadanlıq"])
    with tab_plan:
        render_block1()
        render_block2()
        render_block3()
        render_block4()
        render_block5()
    with tab_inv:
        render_inventory()


if __name__ == "__main__":
    main()
