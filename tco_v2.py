

from dataclasses import dataclass
from typing import List, Dict, Iterable, Optional, Tuple

# ===============================
# Helpers (ported from VBA logic)
# ===============================

def pv_factor(rate: float, year: int) -> float:
    """Present value factor for a given discount rate and year index (>=0)."""
    if year <= 0:
        return 1.0
    return 1.0 / ((1.0 + float(rate)) ** int(year))

def clamp01(x: float) -> float:
    """Clamp a scalar into [0, 1]."""
    return max(0.0, min(1.0, float(x)))

def to_num(v, treat_percent: bool=False) -> float:
    """
    Best-effort conversion to float.
    - Accepts strings with ',' decimal or thousand separators.
    - If treat_percent=True and the value looks like a percent (e.g., '35%' or 0.35),
      returns the decimal fraction (0.35 for 35%).
    """
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        x = float(v)
    else:
        s = str(v).strip()
        # Finnish style decimals: if there's a comma but no dot, treat comma as decimal sep
        if ("," in s) and ("." not in s):
            s = s.replace(" ", "").replace("%", "")
            s = s.replace(".", "")  # kill any thousand dots in weird inputs
            s = s.replace(",", ".")
        else:
            # Remove spaces and thousands-separators
            s = s.replace(" ", "").replace("%", "")
            # If both comma and dot appear, assume dot is decimal and comma is thousands
            if "," in s and "." in s:
                s = s.replace(",", "")
        try:
            x = float(s)
        except ValueError:
            x = 0.0
    if treat_percent:
        # If user typed e.g. 35 -> assume percent; if 0.35, keep as-is
        if abs(x) > 1.0:
            return x / 100.0
        return x
    return x

def to_bool(v) -> bool:
    """Lenient truthiness parser for Excel-like inputs."""
    if isinstance(v, str):
        s = v.strip().lower()
        return s in {"1","true","yes","on","x","taiga","y"}
    return bool(v)


# =============================================
# Cycle table and product price list structures
# =============================================

@dataclass
class CyclePoint:
    year: int
    value_pct: float  # 0..1, i.e., 0.35 for 35%

def cycle_pct(cycle_year: int, cycle_table: Iterable[CyclePoint]) -> float:
    """Find % value for given cycle_year in the cycle table; returns 0.0 if not found."""
    y = int(to_num(cycle_year))
    for row in cycle_table:
        if int(row.year) == y:
            return to_num(row.value_pct, treat_percent=True)
    return 0.0


# =====================
# Core TCO calculations
# =====================

def tco_acquisition_net_tbl(is_taiga, list_price, cycle_year, cycle_table: Iterable[CyclePoint], wacc) -> float:
    """
    Acquisition PV net of discounted buyback (Taiga Cycle).
    VBA signature:
      TCO_AcquisitionNetTbl(isTaiga, listPrice, cycleYear, cycleTable, wacc)
    """
    taiga = to_bool(is_taiga)
    price = to_num(list_price)
    yr = int(to_num(cycle_year))
    r = to_num(wacc)
    if taiga and yr > 0:
        bb_pct = cycle_pct(yr, cycle_table)
        buyback_pv = (price * bb_pct) * pv_factor(r, yr)
    else:
        buyback_pv = 0.0
    # Net acquisition present value = cash-out today (price) minus PV of buyback
    return price - buyback_pv

def tco_operation(is_taiga, area_m2, kwh_m2yr, elec_price, run_frac_trad, occ_rate, standby_taiga, years, wacc) -> float:
    """
    Annual energy cost PV over horizon.
    - If Taiga: run fraction = clamp01(occupancy + standby)
    - Else (traditional): run fraction = clamp01(run_frac_trad)
    Sum PV of (area * kwh/m2/y * run_frac * price) for y = 1..N
    """
    taiga = to_bool(is_taiga)
    A = to_num(area_m2)
    K = to_num(kwh_m2yr)
    p = to_num(elec_price)
    runT = to_num(run_frac_trad)
    occ = to_num(occ_rate)
    stby = to_num(standby_taiga)
    n = int(to_num(years))
    r = to_num(wacc)

    total = 0.0
    for y in range(1, n+1):
        if taiga:
            run_frac = clamp01(occ + stby)  # ~0.40 typical
        else:
            run_frac = clamp01(runT)        # ~0.90 typical
        annual_kwh = K * A * run_frac
        annual_cost = annual_kwh * p
        total += annual_cost * pv_factor(r, y)
    return total

def tco_commissioning(comm_cost, when_year: int=1, wacc: float=0.0) -> float:
    """PV of commissioning cost occurring at a given year (default year 1)."""
    c = to_num(comm_cost)
    y = int(to_num(when_year))
    r = to_num(wacc)
    if y <= 0:
        return c
    return c * pv_factor(r, y)

def tco_maintenance(is_taiga, maint_total, years, wacc) -> float:
    """
    Maintenance cost PV, distributed linearly over the horizon (as in VBA).
    maint_total is the total nominal sum over the horizon; we spread equally per year.
    """
    tot = to_num(maint_total)
    n = int(to_num(years))
    r = to_num(wacc)
    if n <= 0:
        return 0.0
    per_year = tot / n
    pv = 0.0
    for y in range(1, n+1):
        pv += per_year * pv_factor(r, y)
    return pv

def tco_downtime(rate_per_hour, h_install, h_maint_total, years, wacc) -> float:
    """
    Downtime PV: one-time installation hours + maintenance hours distributed linearly.
    """
    rate = to_num(rate_per_hour)
    hI = to_num(h_install)
    hM = to_num(h_maint_total)
    n = int(to_num(years))
    r = to_num(wacc)

    pv_install = (rate * hI) * pv_factor(r, 1) if n >= 1 else (rate * hI)
    pv_maint = 0.0
    if n > 0 and hM > 0:
        per_year = (rate * hM) / n
        for y in range(1, n+1):
            pv_maint += per_year * pv_factor(r, y)
    return pv_install + pv_maint

def tco_end_of_life(is_taiga, eol_cost, years, wacc) -> float:
    """PV of end-of-life cost realized at year N."""
    c = to_num(eol_cost)
    n = int(to_num(years))
    r = to_num(wacc)
    if n <= 0:
        return 0.0
    return c * pv_factor(r, n)

def tco_workspace(tco_pv, space_efficiency_index) -> float:
    """Workspace index = TCO PV / space efficiency index (guard divide-by-zero)."""
    t = to_num(tco_pv)
    s = to_num(space_efficiency_index)
    if s <= 0.0:
        return float("inf")
    return t / s


# =======================
# Aggregate / Yearly view
# =======================

@dataclass
class TCOInputs:
    is_taiga: bool
    years: int
    wacc: float  # as decimal, e.g. 0.08 for 8%

    # Acquisition
    list_price: float

    # Operation
    area_m2: float
    kwh_m2yr: float
    elec_price: float
    run_frac_trad: float   # e.g., 0.90
    occ_rate: float        # e.g., 0.35
    standby_taiga: float   # e.g., 0.05

    # Commissioning
    commissioning_cost: float
    commissioning_year: int

    # Maintenance
    maint_total: float

    # Downtime
    downtime_rate_per_hour: float
    downtime_hours_install: float
    downtime_hours_maint_total: float

    # End of life
    eol_cost: float

    # Taiga Cycle (% table)
    cycle_table: List[CyclePoint]
    cycle_year: int

def compute_tco(inputs: TCOInputs) -> Dict[str, float]:
    """Compute component PVs and total PV."""
    acq = tco_acquisition_net_tbl(inputs.is_taiga, inputs.list_price,
                                  inputs.cycle_year, inputs.cycle_table, inputs.wacc)
    op  = tco_operation(inputs.is_taiga, inputs.area_m2, inputs.kwh_m2yr, inputs.elec_price,
                        inputs.run_frac_trad, inputs.occ_rate, inputs.standby_taiga,
                        inputs.years, inputs.wacc)
    com = tco_commissioning(inputs.commissioning_cost, inputs.commissioning_year, inputs.wacc)
    mai = tco_maintenance(inputs.is_taiga, inputs.maint_total, inputs.years, inputs.wacc)
    dow = tco_downtime(inputs.downtime_rate_per_hour, inputs.downtime_hours_install,
                       inputs.downtime_hours_maint_total, inputs.years, inputs.wacc)
    eol = tco_end_of_life(inputs.is_taiga, inputs.eol_cost, inputs.years, inputs.wacc)

    total = acq + op + com + mai + dow + eol
    return {
        "acquisition_pv": acq,
        "operation_pv": op,
        "commissioning_pv": com,
        "maintenance_pv": mai,
        "downtime_pv": dow,
        "end_of_life_pv": eol,
        "total_pv": total,
    }


# ======================
# Product price handling
# ======================

@dataclass
class Product:
    code: str
    name: str
    unit_price: float

def total_product_price(products: Iterable[Product], selected_codes: Iterable[str]) -> float:
    """
    Sum the unit_price for all selected_codes.
    If duplicates are allowed, pass selected_codes with repetition.
    """
    prices: Dict[str, float] = {p.code: to_num(p.unit_price) for p in products}
    s = 0.0
    for c in selected_codes:
        s += prices.get(c, 0.0)
    return s


# =====================
# Year-by-year breakdown
# =====================

def yearly_breakdown(inputs: TCOInputs) -> List[Dict[str, float]]:
    """
    Return a list of dict rows: year, pv_factor, energy_cost_nominal, energy_cost_pv, maint_pv, downtime_pv, commissioning_pv, eol_pv
    Note: Acquisition is assumed at year 0 (up-front) minus PV(buyback) at cycle year.
    """
    rows: List[Dict[str, float]] = []
    r = inputs.wacc
    n = inputs.years

    # Common run fraction logic
    if inputs.is_taiga:
        run_frac = clamp01(inputs.occ_rate + inputs.standby_taiga)
    else:
        run_frac = clamp01(inputs.run_frac_trad)

    # Per-year components
    energy_nominal = inputs.area_m2 * inputs.kwh_m2yr * run_frac * inputs.elec_price
    maint_per_year = inputs.maint_total / n if n > 0 else 0.0
    dt_per_year = (inputs.downtime_rate_per_hour * inputs.downtime_hours_maint_total) / n if n > 0 else 0.0

    for y in range(1, n+1):
        row = {
            "year": y,
            "pv_factor": pv_factor(r, y),
            "energy_cost_nominal": energy_nominal,
            "energy_cost_pv": energy_nominal * pv_factor(r, y),
            "maintenance_pv": maint_per_year * pv_factor(r, y),
            "downtime_pv": dt_per_year * pv_factor(r, y),
            "commissioning_pv": 0.0,
            "eol_pv": 0.0,
        }
        if y == int(inputs.commissioning_year):
            row["commissioning_pv"] = inputs.commissioning_cost * pv_factor(r, y)
        if y == int(inputs.years):
            row["eol_pv"] = inputs.eol_cost * pv_factor(r, y)
        rows.append(row)

    return rows
