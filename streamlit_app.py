# streamlit_app.py
# ---------------------------------------------------------
# Taiga vs TRAD TCO – per-unit scaling, TRAD €/m² investment,
# TRAD costs scaled by number of rooms (qty) and years.
# ---------------------------------------------------------

import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from tco_v2 import TCOInputs, CyclePoint, compute_tco, yearly_breakdown
from proposal_doc import generate_proposal_doc # make doc available for download

st.set_page_config(page_title="Taiga vs TRAD TCO", layout="wide")

# -------------------- Helpers --------------------

def init_state():
    ss = st.session_state

    # Shared
    ss.setdefault("years", 5)
    ss.setdefault("wacc", 0.05)
    ss.setdefault("area_m2", 4.0)
    ss.setdefault("kwh_m2yr", 105.0)
    ss.setdefault("elec_price", 0.15)
    ss.setdefault("cycle_year", 5)             # Taiga only
    ss.setdefault("area_from_products", None)  # computed from selector

    # Taiga cycle table (only for Taiga)
    if "cycle_df" not in ss:
        ss.cycle_df = pd.DataFrame({
            "year": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
            "value_pct": [50, 50, 50, 40, 35, 30, 25, 20, 15, 10],
        })

    # ---------- Taiga per-unit defaults ----------
    ss.setdefault("taiga_list_price", 10_000.0)
    ss.setdefault("taiga_occ_rate", 0.35)
    ss.setdefault("taiga_standby", 0.05)

    ss.setdefault("taiga_commissioning_cost_unit", 950.0)  # €/unit
    ss.setdefault("taiga_maint_total_unit", 150.0)           # €/unit annually
    ss.setdefault("taiga_dt_rate", 15.0)                    # €/h
    ss.setdefault("taiga_dt_install_h_unit", 8.0)           # h/unit
    ss.setdefault("taiga_dt_maint_h_total_unit", 1.5)       # h/unit annually

    ss.setdefault("taiga_commissioning_year", 1)
    ss.setdefault("taiga_eol_cost", 0.0)                # total (not scaled)

    ss.setdefault("taiga_total_qty", 0)                     # total selected qty

    # ---------- TRAD defaults (per-room / % semantics) ----------
    ss.setdefault("trad_list_price", 10_000.0)              # not used in calc
    ss.setdefault("trad_price_per_m2", 2500.0)              # investment €/m²

    ss.setdefault("trad_commissioning_cost_unit", 1500.0)     # €/room/year
    ss.setdefault("trad_commissioning_year", 1)             # kept for API compatibility
    ss.setdefault("trad_maint_total_unit", 500.0)        # €/room over annually

    ss.setdefault("trad_dt_rate", 15.0)                     # €/h
    ss.setdefault("trad_dt_install_h_unit", 80.0)           # h/room
    ss.setdefault("trad_dt_maint_h_total_unit", 5.0)       # h/room over horizon

    ss.setdefault("trad_eol_pct", 0.20)                     # 20% of investment
    ss.setdefault("trad_run_frac", 0.90)


    # Taiga product selector state
    ss.setdefault("override_price", False)
    ss.setdefault("override_area", False)
    if "price_df" not in ss:
        ss.price_df = pd.DataFrame({
            "code": ["LB1", "LB2", "LB3", "LB5", "LB7", "PIC", "FL10", "FL12", "FL14", "FL21", "FL25", "FL28"],
            "name": ["Taiga LB1", "Taiga LB2", "Taiga LB3", "Taiga LB5", "Taiga LB7", "Picea", "Flex10", "Flex12", "Flex14", "Flex21", "Flex25", "Flex28"],
            "unit_price_eur": [9900, 14900, 15900, 21900, 25900, 18900, 44540, 50290, 57520, 67060, 74080, 82220],
            "area_m2": [1, 2, 3, 5, 7, 3, 10, 12, 14, 21, 25, 28],
            "qty": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        })

def taiga_price_list_ui():
    """Upload/select Taiga products, edit qty; compute total price, total area, total qty."""
    st.markdown("**Taiga product selector (optional)**")
    st.caption("Upload a price list or use the default list. Edit quantities to compute price and area.")

    up = st.file_uploader("Upload price list (xlsx/csv)", type=["xlsx", "csv"], key="pl_up")
    if up is not None:
        try:
            if up.name.lower().endswith(".xlsx"):
                df = pd.read_excel(up)
            else:
                df = pd.read_csv(up)
            cols_map = {c.lower(): c for c in df.columns}
            required = {"code", "name", "unit_price_eur"}
            if not required.issubset(set(cols_map.keys())):
                st.error("File must contain: code, name, unit_price_eur (optional area_m2).")
            else:
                order = ["code", "name", "unit_price_eur"] + (["area_m2"] if "area_m2" in cols_map else [])
                df = df[[cols_map[c] for c in order]]
                if "qty" not in df.columns:
                    df["qty"] = 0
                st.session_state.price_df = df.copy()
        except Exception as e:
            st.error(f"Failed to read file: {e}")

    edited = st.data_editor(
        st.session_state.price_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={"qty": st.column_config.NumberColumn("qty", min_value=0, step=1)},
        key="price_table"
    )

    # Totals (with fillna for safety)
    if not edited.empty:
        unit = edited["unit_price_eur"].fillna(0)
        qty  = edited["qty"].fillna(0)
        total_price = float((unit * qty).sum())
        total_qty   = int(qty.astype(int).sum())
    else:
        total_price, total_qty = 0.0, 0
    st.session_state.taiga_total_qty = total_qty

    st.metric("Computed Taiga list price (€)", f"{total_price:,.0f}")

    total_area = None
    if not edited.empty and "area_m2" in edited.columns:
        try:
            total_area = float((edited["area_m2"].fillna(0) * edited["qty"].fillna(0)).sum())
        except Exception:
            total_area = 0.0

    qcol1, qcol2 = st.columns(2)
    with qcol1:
        st.metric("Selected Taiga quantity (units)", f"{total_qty}")
    with qcol2:
        if total_area is not None:
            st.metric("Area from products (m²)", f"{total_area:,.2f}")

    c1, c2 = st.columns(2)
    with c1:
        st.checkbox("Compute Taiga list price from selected products", key="override_price")
    with c2:
        if total_area is not None:
            st.checkbox("Override shared Area (m²) with total from products", key="override_area")

    if st.session_state.override_price:
        st.session_state.taiga_list_price = total_price

    st.session_state.area_from_products = total_area if total_area is not None else None

    if st.session_state.get("override_area") and st.session_state.area_from_products is not None:
        st.info(f"Using area from products: {st.session_state.area_from_products:.2f} m²", icon="ℹ️")

def to_cycle_list(df: pd.DataFrame):
    """Convert the editable DataFrame to list[CyclePoint]. Accept % or decimals."""
    pts = []
    for r in df.itertuples(index=False):
        year = int(getattr(r, "year"))
        raw = getattr(r, "value_pct")
        try:
            v = float(raw)
            v = v / 100.0 if abs(v) > 1.0 else v
        except Exception:
            v = 0.0
        pts.append(CyclePoint(year=year, value_pct=v))
    return pts

def build_inputs_taiga(shared):
    """Build TCOInputs for Taiga using per-unit fields scaled by qty."""
    ss = st.session_state
    qty = int(ss.get("taiga_total_qty", 0))

    commissioning_total = float(ss.taiga_commissioning_cost_unit) * qty
    maint_total = float(ss.taiga_maint_total_unit) * qty
    dt_install_total_h = float(ss.taiga_dt_install_h_unit) * qty
    dt_maint_total_h   = float(ss.taiga_dt_maint_h_total_unit) * qty

    return TCOInputs(
        is_taiga=True,
        years=int(shared["years"]),
        wacc=float(shared["wacc"]),
        list_price=float(ss.taiga_list_price),
        area_m2=float(shared["area_m2"]),
        kwh_m2yr=float(shared["kwh_m2yr"]),
        elec_price=float(shared["elec_price"]),
        run_frac_trad=0.0,
        occ_rate=float(ss.taiga_occ_rate),
        standby_taiga=float(ss.taiga_standby),
        commissioning_cost=commissioning_total,
        commissioning_year=int(ss.taiga_commissioning_year),
        maint_total=maint_total,
        downtime_rate_per_hour=float(ss.taiga_dt_rate),
        downtime_hours_install=dt_install_total_h,
        downtime_hours_maint_total=dt_maint_total_h,
        eol_cost=float(ss.taiga_eol_cost),
        cycle_table=to_cycle_list(st.session_state.cycle_df),
        cycle_year=int(shared["cycle_year"]),
    )

def build_inputs_trad(shared):
    """Build TCOInputs for TRAD. Costs scaled by rooms (qty) and years where specified."""
    ss = st.session_state
    qty_rooms = int(ss.get("taiga_total_qty", 0))  # rooms = Taiga qty (can be separated later)
    run_frac_trad=float(st.session_state.get("trad_run_frac", 0.90)),

    # Investment = €/m² × area
    list_price_total = float(ss.trad_price_per_m2) * float(shared["area_m2"])

    # Commissioning: €/room/year × rooms × years
    commissioning_total = float(ss.trad_commissioning_cost_unit) * qty_rooms

    # Maintenance: €/room over horizon × rooms
    maint_total = float(ss.trad_maint_total_unit) * qty_rooms

    # Downtime hours (scaled by rooms); rate €/h unchanged
    dt_install_total_h = float(ss.trad_dt_install_h_unit) * qty_rooms
    dt_maint_total_h   = float(ss.trad_dt_maint_h_total_unit) * qty_rooms

    # End-of-life: % of investment
    eol_cost = float(ss.trad_eol_pct) * list_price_total

    return TCOInputs(
        is_taiga=False,
        years=int(shared["years"]),
        wacc=float(shared["wacc"]),
        list_price=list_price_total,
        area_m2=float(shared["area_m2"]),
        kwh_m2yr=float(shared["kwh_m2yr"]),
        elec_price=float(shared["elec_price"]),
        run_frac_trad=float(ss.get("trad_run_frac", 0.90)),
        occ_rate=0.0,
        standby_taiga=0.0,
        commissioning_cost=commissioning_total,
        commissioning_year=int(ss.trad_commissioning_year),
        maint_total=maint_total,
        downtime_rate_per_hour=float(ss.trad_dt_rate),
        downtime_hours_install=dt_install_total_h,
        downtime_hours_maint_total=dt_maint_total_h,
        eol_cost=eol_cost,
        cycle_table=[],
        cycle_year=0,
    )

def excel_bytes_all(taiga_summary, trad_summary, df_taiga, df_trad, df_delta, taiga_inp, trad_inp):
    """Build an Excel workbook with all outputs."""
    from openpyxl import Workbook
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Summary-Taiga"
    ws1.append(["Component", "PV (€)"])
    for k, v in taiga_summary.items():
        ws1.append([k, float(v)])

    ws2 = wb.create_sheet("Summary-TRAD")
    ws2.append(["Component", "PV (€)"])
    for k, v in trad_summary.items():
        ws2.append([k, float(v)])

    ws3 = wb.create_sheet("Summary-Delta (T-TR)")
    all_keys = sorted(set(taiga_summary.keys()) | set(trad_summary.keys()))
    ws3.append(["Component", "Delta PV (€)"])
    for k in all_keys:
        ws3.append([k, float(taiga_summary.get(k, 0.0) - trad_summary.get(k, 0.0))])

    for title, df in [("Yearly-Taiga", df_taiga), ("Yearly-TRAD", df_trad), ("Yearly-Delta", df_delta)]:
        ws = wb.create_sheet(title)
        ws.append(list(df.columns))
        for _, r in df.iterrows():
            ws.append(list(r.values))

    ws4 = wb.create_sheet("CycleTable (Taiga only)")
    ws4.append(["Year", "Value%"])
    for cp in taiga_inp.cycle_table:
        ws4.append([cp.year, cp.value_pct])

    for title, inp in [("Inputs-Taiga", taiga_inp), ("Inputs-TRAD", trad_inp)]:
        ws = wb.create_sheet(title)
        for k, v in inp.__dict__.items():
            if k == "cycle_table":
                continue
            ws.append([k, v])

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

def _total_pv(summary: dict) -> float:
    """Return overall Total PV from a summary dict."""
    if "total_pv" in summary:
        try:
            return float(summary["total_pv"])
        except Exception:
            pass
    return float(sum(v for v in summary.values() if isinstance(v, (int, float))))

# -------------------- UI --------------------

init_state()
st.title("Taiga vs Traditional Building — Total Cost of Ownership")

with st.expander("Shared parameters", expanded=False):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.number_input("Horizon (years)", 1, 50, key="years")
    with c2:
        st.number_input("WACC (decimal, e.g., 0.08)", 0.0, 1.0, key="wacc", step=0.01)
    with c3:
        st.number_input("Area (m²)", 0.0, 1e9, key="area_m2", step=1.0)
    with c4:
        st.number_input("kWh / m² / year", 0.0, 1e9, key="kwh_m2yr", step=1.0)
    c5, c6 = st.columns(2)
    with c5:
        st.number_input("Electricity price (€/kWh)", 0.0, 10.0, key="elec_price", step=0.01)
    with c6:
        st.caption("Cycle settings are under Taiga section.")

tab_products = st.tabs(["Products"])[0]
with tab_products:
    st.markdown("### Products and Parameters")
    colT, colR = st.columns(2, gap="large")

    with colT:
        st.subheader("Taiga")

        with st.expander("Taiga Pricing", expanded=False):

            taiga_price_list_ui()

        with st.expander("Taiga Parameters", expanded=False):

            st.number_input(
                "List price (€)",
                0.0, 1e12,
                key="taiga_list_price",
                step=1000.0,
                disabled=st.session_state.get("override_price", False)
            )

            st.slider("Occupancy rate", 0.0, 1.0, key="taiga_occ_rate")
            st.slider("Standby share", 0.0, 1.0, key="taiga_standby")

            st.number_input("Commissioning cost per unit (€)", 0.0, 1e12,
                            key="taiga_commissioning_cost_unit", step=50.0)
            st.number_input("Maintenance per unit annually (€)", 0.0, 1e12,
                            key="taiga_maint_total_unit", step=10.0)

            c1, c2, c3 = st.columns(3)
            with c1:
                st.number_input("Downtime rate (€/h)", 0.0, 1e9, key="taiga_dt_rate", step=1.0)
            with c2:
                st.number_input("Install downtime per unit (h)", 0.0, 1e9,
                                key="taiga_dt_install_h_unit", step=0.5)
            with c3:
                st.number_input("Maint downtime per unit annually (h)", 0.0, 1e9,
                                key="taiga_dt_maint_h_total_unit", step=0.5)
            st.number_input("Commissioning year", 0, 50, key="taiga_commissioning_year")
            st.number_input("End-of-life cost (total) (€)", 0.0, 1e12, key="taiga_eol_cost", step=100.0)

        qty = st.session_state.get("taiga_total_qty", 0)
        
        d1, d2, d3, d4 = st.columns(4)
        with d1:
            st.metric("Commissioning total (€)",
                      f"{st.session_state.taiga_commissioning_cost_unit * qty:,.0f}")
        with d2:
            st.metric("Maintenance total (€)",
                      f"{st.session_state.taiga_maint_total_unit * qty * int(st.session_state.years):,.0f}")
        with d3:
            st.metric("Install downtime total (h)",
                      f"{st.session_state.taiga_dt_install_h_unit * qty:,.1f}")
        with d4:
            st.metric("Maint downtime total (h)",
                      f"{st.session_state.taiga_dt_maint_h_total_unit * qty * int(st.session_state.years):,.0f}")

        with st.expander("Taiga Buyback", expanded=False):
            st.markdown("### Taiga buyback (Cycle)")
            st.number_input("Cycle (buyback) year (Taiga)", 0, 50, key="cycle_year")
            st.markdown("**Taiga Cycle table (Year, Value %)**")
            st.caption("Edit or add rows; values can be 70 (%=0.70) or 0.70.")
            st.session_state.cycle_df = st.data_editor(
                st.session_state.cycle_df,
                num_rows="dynamic",
                use_container_width=True
            )  # <-- ensure this closing parenthesis exists

    # Compute effective area after Taiga selector potentially changed area_from_products
    effective_area = (
        st.session_state.area_from_products
        if (st.session_state.get("override_area") and st.session_state.get("area_from_products") is not None)
        else st.session_state.area_m2
    )

    with colR:
        st.subheader("Traditional Building (TRAD)")

        with st.expander("Building Pricing", expanded=False):

            st.number_input("Investment price per m² (TRAD) (€)",
                            0.0, 1e6, key="trad_price_per_m2", step=50.0)
        
        with st.expander("Building Parameters", expanded=False):
            st.slider("Run fraction (Traditional)", 0.0, 1.0, key="trad_run_frac")

            # End-of-life % input (store as fraction in session)
            eol_pct_input = st.number_input("End-of-life (% of investment)", 0.0, 100.0,
                                            value=st.session_state.trad_eol_pct * 100, step=1.0)
            st.session_state.trad_eol_pct = float(eol_pct_input) / 100.0

            # Per-room inputs for TRAD
            st.number_input("Commissioning cost per room (€)", 0.0, 1e9,
                            key="trad_commissioning_cost_unit", step=10.0)
            st.number_input("Maintenance total per room annually (€)", 0.0, 1e12,
                            key="trad_maint_total_unit", step=100.0)

            c1, c2, c3 = st.columns(3)
            with c1:
                st.number_input("Downtime rate (€/h)", 0.0, 1e9, key="trad_dt_rate", step=1.0)
            with c2:
                st.number_input("Install downtime per room (h)", 0.0, 1e9,
                                key="trad_dt_install_h_unit", step=1.0)
            with c3:
                st.number_input("Maint downtime per room annually (h)", 0.0, 1e9,
                                key="trad_dt_maint_h_total_unit", step=1.0)

            st.number_input("Commissioning year (kept for API)", 0, 50, key="trad_commissioning_year")

        # Show derived TRAD totals
        qty_rooms = st.session_state.get("taiga_total_qty", 0)
        trad_invest_total = float(st.session_state.trad_price_per_m2) * float(effective_area)
        trad_comm_total   = float(st.session_state.trad_commissioning_cost_unit) * qty_rooms
        trad_maint_total  = float(st.session_state.trad_maint_total_unit) * qty_rooms * int(st.session_state.years)
        trad_dt_install_h_total = float(st.session_state.trad_dt_install_h_unit) * qty_rooms
        trad_dt_maint_h_total   = float(st.session_state.trad_dt_maint_h_total_unit) * qty_rooms * int(st.session_state.years)
        trad_eol_total = float(st.session_state.trad_eol_pct) * trad_invest_total

        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("Computed TRAD investment (€)", f"{trad_invest_total:,.0f}")
        with m2:
            st.metric("Rooms (qty) used", f"{qty_rooms}")
        with m3:
            st.metric("Years", f"{int(st.session_state.years)}")

        m4, m5, m6, m7 = st.columns(4)
        with m4:
            st.metric("Commissioning total (€)", f"{trad_comm_total:,.0f}")
        with m5:
            st.metric("Maintenance total (€)", f"{trad_maint_total:,.0f}")
        with m6:
            st.metric("Install downtime total (h)", f"{trad_dt_install_h_total:,.1f}")
        with m7:
            st.metric("Maint downtime total (h)", f"{trad_dt_maint_h_total:,.1f}")

        st.metric("End-of-life total (€)", f"{trad_eol_total:,.0f}")

# Area info line
st.caption(
    f"Using area for calculations: "
    f"{(st.session_state.area_from_products if st.session_state.get('override_area') and st.session_state.get('area_from_products') is not None else st.session_state.area_m2):.2f} m²"
)

# Build inputs and compute
shared = dict(
    years=st.session_state.years,
    wacc=st.session_state.wacc,
    area_m2=effective_area,
    kwh_m2yr=st.session_state.kwh_m2yr,
    elec_price=st.session_state.elec_price,
    cycle_year=st.session_state.cycle_year,
)
taiga_inp = build_inputs_taiga(shared)
trad_inp  = build_inputs_trad(shared)

taiga_sum = dict(compute_tco(taiga_inp))
trad_sum  = dict(compute_tco(trad_inp))

taiga_rows = yearly_breakdown(taiga_inp)
trad_rows  = yearly_breakdown(trad_inp)

df_taiga = pd.DataFrame(taiga_rows)
df_trad  = pd.DataFrame(trad_rows)

# Delta
df_delta = None
if not df_taiga.empty and not df_trad.empty:
    common_cols = [c for c in df_taiga.columns if c in df_trad.columns]
    sum_cols = [c for c in common_cols if c != "year"]
    merged = df_taiga[["year"] + sum_cols].merge(
        df_trad[["year"] + sum_cols], on="year", suffixes=("_T", "_TR")
    )
    for c in sum_cols:
        merged[c] = merged[f"{c}_T"] - merged[f"{c}_TR"]
    df_delta = merged[["year"] + sum_cols]

with st.expander("Results", expanded=False):
    st.markdown("---")
    st.subheader("Summary (Present Value €) — live")

    cols = st.columns(3)
    with cols[0]:
        st.caption("Taiga")
        st.dataframe(pd.DataFrame([taiga_sum]), use_container_width=True, hide_index=True)
    with cols[1]:
        st.caption("TRAD")
        st.dataframe(pd.DataFrame([trad_sum]), use_container_width=True, hide_index=True)
    with cols[2]:
        st.caption("Delta (Taiga − TRAD)")
        keys = sorted(set(taiga_sum.keys()) | set(trad_sum.keys()))
        delta = {k: taiga_sum.get(k, 0.0) - trad_sum.get(k, 0.0) for k in keys}
        st.dataframe(pd.DataFrame([delta]), use_container_width=True, hide_index=True)

    st.markdown("### Yearly breakdowns")
    colA, colB, colC = st.columns(3)
    with colA:
        st.caption("Taiga")
        st.dataframe(df_taiga, use_container_width=True, hide_index=True)
    with colB:
        st.caption("TRAD")
        st.dataframe(df_trad, use_container_width=True, hide_index=True)
    with colC:
        st.caption("Delta (Taiga − TRAD)")
        if df_delta is not None:
            st.dataframe(df_delta, use_container_width=True, hide_index=True)
        else:
            st.info("Delta becomes available when both sides have rows.", icon="ℹ️")

# Overall Total PV metrics
taiga_total_pv = _total_pv(taiga_sum)
trad_total_pv  = _total_pv(trad_sum)
delta_total_pv = taiga_total_pv - trad_total_pv

st.markdown("### Overall Total PV")
m1, m2, m3 = st.columns(3)
with m1:
    st.metric("Taiga Total PV (€)", f"{taiga_total_pv:,.0f}")
with m2:
    st.metric("TRAD Total PV (€)", f"{trad_total_pv:,.0f}")
with m3:
    st.metric("Delta Total PV (€)  (Taiga − TRAD)", f"{delta_total_pv:,.0f}",
              delta=f"{delta_total_pv:,.0f}")

# Export
st.markdown("### Export")
if st.button("Download Excel (Taiga, TRAD, Delta)"):
    xbytes = excel_bytes_all(
        taiga_sum, trad_sum,
        df_taiga, df_trad,
        (df_delta if df_delta is not None else pd.DataFrame()),
        taiga_inp, trad_inp
    )
    st.download_button(
        "Save TCO_Comparison.xlsx",
        data=xbytes,
        file_name="TCO_Comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- Export Word ---
st.markdown("### Export to Word")
payload = {
    "customer_name": "Demo Customer",
    "project_name": "Demo Project",
    "date_str": pd.Timestamp.now().strftime("%Y-%m-%d"),
    "params": shared,
    "results": {
        "TCO_TRAD_PV": trad_total_pv,
        "TCO_TAIGA_PV": taiga_total_pv,
        "DIFF_TRAD_TAIGA": trad_total_pv - taiga_total_pv,
    },
}
doc_bytes = generate_proposal_doc(payload, locale="fi_FI")

st.download_button(
    label="📄 Download Word Summary",
    data=doc_bytes,
    file_name="TCO_Summary.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)


# Cost comparison chart
with st.expander("Chart", expanded=False):
    st.markdown("### Cost comparison chart (PV per year)")
    required_cols = ["energy_cost_pv", "maintenance_pv", "downtime_pv", "commissioning_pv", "eol_pv"]
    have_cols_taiga = (not df_taiga.empty) and all(c in df_taiga.columns for c in required_cols)
    have_cols_trad  = (not df_trad.empty)  and all(c in df_trad.columns  for c in required_cols)

    if have_cols_taiga and have_cols_trad:
        taiga_total = df_taiga[required_cols].sum(axis=1)
        trad_total  = df_trad[required_cols].sum(axis=1)
        years = df_taiga["year"]

        fig, ax = plt.subplots(figsize=(10, 5), dpi=100)
        ax.plot(years, taiga_total, label="Taiga")
        ax.plot(years, trad_total, label="TRAD")
        ax.plot(years, taiga_total - trad_total, label="Delta (Taiga − TRAD)")
        ax.set_xlabel("Year")
        ax.set_ylabel("PV cost (€)")
        ax.set_title("Yearly PV cost comparison")
        ax.legend()
        ax.grid(True, which="both", linewidth=0.5, alpha=0.5)
        st.pyplot(fig)
    else:
        st.info("Chart becomes available when both Taiga and TRAD yearly tables have data.", icon="ℹ️")
