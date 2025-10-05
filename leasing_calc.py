# leasing_calc.py
from __future__ import annotations
from dataclasses import dataclass
from typing import List, Optional, Tuple
import math
import pandas as pd

# -------------------- Data structures --------------------

@dataclass
class CyclePoint:
    year: int
    value_pct: float  # e.g. 0.30 for 30%

@dataclass
class LeasingInputs:
    list_price: float                 # Taigan kokonaishinta (€)
    wacc_annual: float               # vuosikorko desimaalina, esim. 0.08
    term_years: int                  # leasing-aika vuosina (esim. 5)
    cycle_year: int = 0              # buyback-vuosi (jos 0 → ei buyback)
    cycle_table: Optional[List[CyclePoint]] = None

# -------------------- Helpers --------------------

def _norm_factor(x) -> float:
    """Hyväksyy 1.8 tai 0.018 → palauttaa aina desimaalina/kk (0.018)."""
    try:
        v = float(x)
    except Exception:
        return 0.0
    return v/100.0 if abs(v) > 1.0 else v

def monthly_rate_from_annual(wacc_annual: float) -> float:
    """Kuukausikorko WACC:sta (diskonttaus)."""
    return (1.0 + float(wacc_annual))**(1.0/12.0) - 1.0

def buyback_pct_for_year(cycle_year: int, cycle_table: Optional[List[CyclePoint]]) -> float:
    if not cycle_table or cycle_year <= 0:
        return 0.0
    for cp in cycle_table:
        if int(cp.year) == int(cycle_year):
            v = float(cp.value_pct)
            return v/100.0 if abs(v) > 1.0 else v
    return 0.0

# -------------------- Core computations --------------------

def monthly_factor_for_term(term_years: int, factor_table: pd.DataFrame) -> float:
    """
    Palauttaa kuukaudessa maksettavan kerroin-arvon (desimaalina/kk) annetulle termille.
    Taulukossa voi olla sarakkeet:
      - ("term_years", "monthly_factor") TAI
      - ("term_months", "monthly_factor")
    'monthly_factor' saa olla % tai desimaali.
    Jos täsmällistä termiä ei löydy, käytetään lähintä pienempää; ellei sitäkään, pienintä.
    """
    if factor_table is None or factor_table.empty:
        return 0.0

    cols = {c.lower(): c for c in factor_table.columns}
    if "term_years" in cols:
        df = factor_table[[cols["term_years"], cols.get("monthly_factor", "monthly_factor")]].copy()
        df.columns = ["term_years", "monthly_factor"]
        df["monthly_factor"] = df["monthly_factor"].apply(_norm_factor)
        # etsi täsmäys tai lähin pienempi
        cand = df.loc[df["term_years"] == term_years]
        if not cand.empty:
            return float(cand["monthly_factor"].iloc[0])
        # lähin pienempi
        smaller = df.loc[df["term_years"] < term_years].sort_values("term_years")
        if not smaller.empty:
            return float(smaller["monthly_factor"].iloc[-1])
        # fallback: pienin
        return float(df.sort_values("term_years").iloc[0]["monthly_factor"])

    if "term_months" in cols:
        df = factor_table[[cols["term_months"], cols.get("monthly_factor", "monthly_factor")]].copy()
        df.columns = ["term_months", "monthly_factor"]
        df["monthly_factor"] = df["monthly_factor"].apply(_norm_factor)
        target_mo = int(term_years) * 12
        cand = df.loc[df["term_months"] == target_mo]
        if not cand.empty:
            return float(cand["monthly_factor"].iloc[0])
        smaller = df.loc[df["term_months"] < target_mo].sort_values("term_months")
        if not smaller.empty:
            return float(smaller["monthly_factor"].iloc[-1])
        return float(df.sort_values("term_months").iloc[0]["monthly_factor"])

    # ei sopivaa termikolumnia
    return 0.0

def monthly_payment_base(list_price: float, monthly_factor: float) -> float:
    """Perus kk-maksu ilman buybackia."""
    return float(list_price) * float(monthly_factor)

def equivalent_level_monthly_from_lump(lump: float, months_to_lump: int, monthly_rate: float, total_months: int) -> float:
    """
    Muuntaa tulevaisuuden kertasumman (lump @ month M) tasaiseksi kk-sarjaksi
    koko N kuukauden ajalle (1..N).
    - Lasketaan ensin PV(0) tuolle kertasummalle.
    - Sitten jaetaan annuiteettitekijällä AF = sum_{t=1..N} 1/(1+r)^t.
    Palautetaan kk-hyvityksen suuruus (positiivinen → vähennetään kk-maksusta).
    """
    if lump == 0 or months_to_lump <= 0 or total_months <= 0:
        return 0.0

    r = float(monthly_rate)
    # PV kertasummalle
    pv_lump = float(lump) / ((1.0 + r) ** months_to_lump)
    # annuiteettitekijä koko kaudelle
    if r == 0:
        af = float(total_months)
    else:
        af = (1.0 - (1.0 + r) ** (-total_months)) / r
    return pv_lump / af

def monthly_payment_with_buyback(
    list_price: float,
    monthly_factor: float,
    wacc_annual: float,
    term_years: int,
    cycle_year: int,
    cycle_table: Optional[List[CyclePoint]],
) -> Tuple[float, float]:
    """
    Palauttaa (base_monthly, monthly_with_buyback_equivalent).
    - base_monthly = list_price * monthly_factor
    - monthly_with_buyback_equivalent = base_monthly - kk-hyvitys, jos buyback osuu kauteen
    """
    base = monthly_payment_base(list_price, monthly_factor)
    months_total = int(term_years) * 12
    if cycle_year <= 0 or cycle_table is None:
        return base, base

    if cycle_year > term_years:
        # Buyback vasta kauden jälkeen → ei vaikutusta kauden kk-erään
        return base, base

    pct = buyback_pct_for_year(cycle_year, cycle_table)
    if pct <= 0:
        return base, base

    lump = float(list_price) * float(pct)  # buyback-kertakorvaus
    r_m = monthly_rate_from_annual(wacc_annual)
    months_to_lump = int(cycle_year) * 12
    credit = equivalent_level_monthly_from_lump(lump, months_to_lump, r_m, months_total)
    return base, max(0.0, base - credit)

# -------------------- Yearly PV table for leasing --------------------

def leasing_yearly_pv_table(
    list_price: float,
    monthly_factor: float,
    wacc_annual: float,
    term_years: int,
    cycle_year: int = 0,
    cycle_table: Optional[List[CyclePoint]] = None,
) -> pd.DataFrame:
    """
    Tuottaa DataFramen, jossa vuodet riveinä (0..term_years) ja sarakkeet:
      - leasing_pv (vuosittain kk-maksujen PV-summa)
      - buyback_pv (negatiivinen PV siinä vuonna, jos buyback osuu)
      - total_pv (leasing_pv + buyback_pv)
    Huom: Year 0 sisältää kk-1..kk-12 kassavirrat diskontattuna.
    """
    months_total = int(term_years) * 12
    r_m = monthly_rate_from_annual(wacc_annual)
    base_mo = monthly_payment_base(list_price, monthly_factor)

    # PV per month, summataan vuosittain
    pv_by_year = {y: 0.0 for y in range(0, int(term_years) + 1)}  # 0..term_years
    m_start = 1
    for y in range(0, int(term_years)):
        m_end = min((y + 1) * 12, months_total)
        pv_sum = 0.0
        for m in range(m_start, m_end + 1):
            pv_sum += base_mo / ((1.0 + r_m) ** m)
        pv_by_year[y] = pv_sum
        m_start = m_end + 1

    # Buyback PV (jos osuu kaudelle)
    buyback_by_year = {y: 0.0 for y in range(0, int(term_years) + 1)}
    if cycle_year > 0 and cycle_year <= term_years:
        pct = buyback_pct_for_year(cycle_year, cycle_table)
        if pct > 0:
            lump = float(list_price) * pct
            months_to_lump = int(cycle_year) * 12
            pv_lump = - lump / ((1.0 + r_m) ** months_to_lump)
            buyback_by_year[int(cycle_year)] = pv_lump

    rows = []
    for y in range(0, int(term_years) + 1):
        leasing_pv = pv_by_year.get(y, 0.0)
        buyback_pv = buyback_by_year.get(y, 0.0)
        rows.append({"year": y, "leasing_pv": leasing_pv, "buyback_pv": buyback_pv,
                     "total_pv": leasing_pv + buyback_pv})
    return pd.DataFrame(rows)

# -------------------- Pivot for display --------------------

def pivot_leasing_for_display(df_yearly: pd.DataFrame) -> pd.DataFrame:
    """
    Muuttaa leasingin vuositaulukon pivotiksi:
      - rivit = ["Leasing (PV)", "Buyback (−)", "— Total (per year) —"]
      - sarakkeet = "Year 0..N" + "Total"
    """
    if df_yearly is None or df_yearly.empty:
        return pd.DataFrame()

    value_cols = ["leasing_pv", "buyback_pv"]
    df = df_yearly[["year"] + value_cols].copy()
    pv = df.set_index("year")[value_cols].T
    pv.index = ["Leasing (PV)", "Buyback (−)"]
    pv.columns = [f"Year {c}" for c in pv.columns]
    pv["Total"] = pv.sum(axis=1)

    # per-year total-rivi
    year_cols = [c for c in pv.columns if c.startswith("Year ")]
    per_year_total = pv[year_cols].sum(axis=0)
    per_year_total["Total"] = per_year_total.sum()
    pv.loc["— Total (per year) —"] = per_year_total
    return pv
