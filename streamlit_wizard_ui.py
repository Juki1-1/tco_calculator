import base64
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from tco_v2 import TCOInputs
import leasing_calc as lc


TAIGA_EXTRA_CSS = """
:root {
  --taiga-black: #0D0D0D;
  --taiga-white: #FAFAF8;
  --taiga-green: #1E4D35;
  --taiga-green-mid: #2E7D52;
  --taiga-green-light: #E8F2EC;
  --taiga-gray-100: #F4F4F0;
  --taiga-gray-200: #E0DFD8;
  --taiga-gray-400: #9E9D96;
  --taiga-gray-600: #5C5B56;
  --taiga-amber: #C47C2A;
  --taiga-red: #A83228;
  --taiga-font-display: 'Georgia', 'Times New Roman', serif;
  --taiga-font-body: 'Helvetica Neue', Arial, sans-serif;
  --taiga-font-mono: 'Courier New', monospace;
  --taiga-radius-sm: 2px;
  --taiga-radius-md: 4px;
  --taiga-radius-lg: 6px;
  --taiga-border: 1px solid var(--taiga-gray-200);
  --taiga-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
}

html, body, [class*="css"] {
  font-family: var(--taiga-font-body);
}

.block-container {
  padding-top: 2rem;
  padding-bottom: 3rem;
}

.taiga-form-section {
  max-width: 520px;
}

.taiga-hero {
  background: linear-gradient(135deg, rgba(30,77,53,0.08), rgba(196,124,42,0.08));
  border: var(--taiga-border);
  border-radius: var(--taiga-radius-lg);
  padding: 1.2rem 1.4rem 1.1rem;
  margin-bottom: 1rem;
}

.taiga-title {
  font-family: var(--taiga-font-display);
  font-size: 2.2rem;
  line-height: 1.05;
  letter-spacing: -0.02em;
  color: var(--taiga-black);
  margin: 0;
}

.taiga-subtitle {
  color: var(--taiga-gray-600);
  margin-top: 0.45rem;
  max-width: 72ch;
  font-size: 0.98rem;
  line-height: 1.45;
}

.taiga-brandbar {
  display: grid;
  grid-template-columns: 110px minmax(0, 1fr);
  gap: 1rem;
  align-items: center;
}

.taiga-brandbar__logo {
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(255,255,255,0.65);
  border: 1px solid rgba(13,13,13,0.06);
  border-radius: 10px;
  min-height: 84px;
  padding: 0.45rem;
}

.taiga-brandbar__logo img {
  max-width: 100%;
  max-height: 72px;
  object-fit: contain;
  filter: contrast(1.02);
}

.taiga-brandbar__copy {
  min-width: 0;
}

.taiga-progress {
  display: grid;
  grid-template-columns: repeat(7, minmax(0, 1fr));
  gap: 0.45rem;
  margin: 0.9rem 0 1rem;
}

.taiga-progress__step {
  min-width: 0;
  border-radius: 999px;
  padding: 0.52rem 0.7rem 0.58rem;
  border: 1px solid var(--taiga-gray-200);
  background: var(--taiga-white);
}

.taiga-progress__step--done {
  background: rgba(30,77,53,0.08);
  border-color: rgba(30,77,53,0.18);
}

.taiga-progress__step--active {
  background: linear-gradient(90deg, var(--taiga-green), var(--taiga-green-mid));
  border-color: var(--taiga-green);
}

.taiga-progress__kicker {
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--taiga-gray-400);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.taiga-progress__label {
  margin-top: 0.12rem;
  color: var(--taiga-black);
  font-size: 0.9rem;
  font-weight: 600;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.taiga-progress__step--active .taiga-progress__kicker,
.taiga-progress__step--active .taiga-progress__label {
  color: var(--taiga-white);
}

.taiga-label {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.1em;
  color: var(--taiga-gray-400);
}

.block-container h3 {
  font-family: var(--taiga-font-display);
  font-size: 1.85rem;
  line-height: 1.1;
  letter-spacing: -0.015em;
  margin-bottom: 0.2rem;
}

.block-container h4 {
  font-size: 1.08rem;
  line-height: 1.2;
  margin-bottom: 0.2rem;
}

[data-testid="stCaptionContainer"] p {
  color: var(--taiga-gray-600);
  font-size: 0.92rem;
  line-height: 1.45;
}

.taiga-card {
  background: var(--taiga-white);
  border: var(--taiga-border);
  border-radius: var(--taiga-radius-md);
  padding: 1.1rem 1.15rem;
  box-shadow: var(--taiga-shadow);
  height: 100%;
  min-height: 152px;
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
}

.taiga-card--dark {
  background: var(--taiga-black);
  color: var(--taiga-white);
}

.taiga-card__title {
  font-size: 0.78rem;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--taiga-gray-400);
  line-height: 1.35;
  min-height: 2.1rem;
}

.taiga-card__value {
  font-family: var(--taiga-font-mono);
  font-size: clamp(1.35rem, 1.1rem + 1vw, 1.75rem);
  margin-top: 0.35rem;
  color: var(--taiga-black);
  line-height: 1.15;
  min-height: 3.1rem;
  display: flex;
  align-items: flex-start;
}

.taiga-card__note {
  margin-top: auto;
  padding-top: 0.45rem;
  color: var(--taiga-gray-600);
  font-size: 0.88rem;
  line-height: 1.35;
  min-height: 2.6rem;
}

.taiga-card--dark .taiga-card__title,
.taiga-card--dark .taiga-card__value,
.taiga-card--dark .taiga-card__note {
  color: var(--taiga-white);
}

.taiga-picker {
  background: var(--taiga-gray-100);
  border: var(--taiga-border);
  border-radius: var(--taiga-radius-md);
  padding: 0.9rem 1rem;
  margin-bottom: 1rem;
}

.taiga-subsection {
  background: rgba(244, 244, 240, 0.55);
  border: var(--taiga-border);
  border-radius: var(--taiga-radius-md);
  padding: 0.95rem 1rem 0.8rem;
  margin-top: 1rem;
}

.taiga-subsection--compact {
  padding: 0.75rem 0.85rem 0.65rem;
  margin-top: 0.8rem;
}

.taiga-subsection__title {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.1em;
  color: var(--taiga-gray-400);
  margin-bottom: 0.3rem;
}

.taiga-subsection__body {
  color: var(--taiga-gray-600);
  font-size: 0.92rem;
  margin-bottom: 0.75rem;
}

.taiga-picker__title {
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--taiga-gray-400);
  margin-bottom: 0.35rem;
}

.taiga-picker__body {
  color: var(--taiga-gray-600);
  font-size: 0.95rem;
}

.taiga-catalog {
  margin-top: 0.9rem;
}

.taiga-family-section + .taiga-family-section {
  margin-top: 1rem;
}

.taiga-family-title {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--taiga-gray-400);
  margin-bottom: 0.5rem;
}

.taiga-catalog-card {
  background: var(--taiga-white);
  border: var(--taiga-border);
  border-radius: var(--taiga-radius-md);
  padding: 0.8rem 0.9rem;
  margin-bottom: 0.55rem;
}

.taiga-catalog-card__head {
  display: flex;
  justify-content: space-between;
  gap: 0.75rem;
  align-items: baseline;
}

.taiga-catalog-card__title {
  color: var(--taiga-black);
  font-size: 0.98rem;
  font-weight: 600;
}

.taiga-catalog-card__meta {
  margin-top: 0.2rem;
  color: var(--taiga-gray-600);
  font-size: 0.88rem;
}

.taiga-catalog-card__status {
  display: inline-flex;
  align-items: center;
  gap: 0.35rem;
  margin-left: 0.35rem;
}

.taiga-catalog-card__status-pill {
  display: inline-flex;
  align-items: center;
  padding: 0.12rem 0.42rem;
  border-radius: 999px;
  font-size: 0.76rem;
  line-height: 1;
  border: 1px solid rgba(224, 223, 216, 0.9);
  background: var(--taiga-gray-100);
  color: var(--taiga-gray-600);
}

.taiga-catalog-card__status-pill--active {
  border-color: rgba(30, 77, 53, 0.14);
  background: var(--taiga-green-light);
  color: var(--taiga-green);
}

.taiga-catalog-card__price {
  font-family: var(--taiga-font-mono);
  font-size: 0.92rem;
  white-space: nowrap;
}

.taiga-line-items {
  margin-top: 1rem;
  background: var(--taiga-white);
  border: var(--taiga-border);
  border-radius: var(--taiga-radius-md);
  padding: 0.9rem 1rem;
}

.taiga-line-items__header {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--taiga-gray-400);
  margin-bottom: 0.6rem;
}

.taiga-line-item {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto auto;
  gap: 0.75rem;
  align-items: baseline;
  padding: 0.55rem 0;
}

.taiga-line-item + .taiga-line-item {
  border-top: 1px solid rgba(224, 223, 216, 0.65);
}

.taiga-line-item__title {
  color: var(--taiga-black);
  font-size: 0.96rem;
}

.taiga-line-item__meta {
  color: var(--taiga-gray-600);
  font-size: 0.86rem;
  margin-top: 0.1rem;
}

.taiga-line-item__qty,
.taiga-line-item__value {
  font-family: var(--taiga-font-mono);
  font-size: 0.95rem;
  color: var(--taiga-black);
  white-space: nowrap;
}

.taiga-family-badge {
  display: inline-block;
  margin-right: 0.45rem;
  padding: 0.12rem 0.45rem;
  border-radius: 999px;
  font-size: 10px;
  letter-spacing: 0.06em;
  text-transform: uppercase;
  background: var(--taiga-green-light);
  color: var(--taiga-green);
  vertical-align: middle;
}

.taiga-selected-pills {
  display: flex;
  flex-wrap: wrap;
  gap: 0.45rem;
  margin: 0.75rem 0 0.9rem;
}

.taiga-selected-pill {
  display: inline-flex;
  align-items: center;
  gap: 0.45rem;
  padding: 0.38rem 0.7rem;
  border-radius: 999px;
  background: rgba(232, 242, 236, 0.85);
  border: 1px solid rgba(30, 77, 53, 0.18);
  color: var(--taiga-green);
  font-size: 0.86rem;
  line-height: 1;
}

.taiga-selected-pill__qty {
  font-family: var(--taiga-font-mono);
  color: var(--taiga-white);
  background: var(--taiga-green-mid);
  border-radius: 999px;
  padding: 0.18rem 0.42rem;
  font-size: 0.78rem;
  font-weight: 700;
  letter-spacing: 0.01em;
}

.taiga-qty-row__title {
  color: var(--taiga-black);
  font-size: 0.94rem;
  font-weight: 600;
}

.taiga-qty-row__meta {
  color: var(--taiga-gray-600);
  font-size: 0.84rem;
  margin-top: 0.08rem;
}

.taiga-qty-row__value {
  font-family: var(--taiga-font-mono);
  color: var(--taiga-black);
  font-size: 0.9rem;
  text-align: right;
  white-space: nowrap;
  margin-top: 1.9rem;
}

.taiga-inline-note {
  color: var(--taiga-gray-600);
  font-size: 0.87rem;
  margin: 0.15rem 0 0.55rem;
}

.taiga-inline-note--warning {
  color: var(--taiga-red);
  font-weight: 600;
}

.taiga-flow-step {
  background: rgba(250, 250, 248, 0.92);
  border: var(--taiga-border);
  border-radius: var(--taiga-radius-md);
  padding: 0.9rem 1rem 0.95rem;
  margin: 0.8rem 0 1rem;
}

.taiga-flow-step--compact {
  padding: 0.72rem 0.82rem 0.75rem;
  margin: 0.55rem 0 0.75rem;
}

.taiga-flow-step__eyebrow {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--taiga-gray-400);
  margin-bottom: 0.25rem;
}

.taiga-flow-step__title {
  color: var(--taiga-black);
  font-size: 1.02rem;
  font-weight: 600;
  margin-bottom: 0.2rem;
  line-height: 1.2;
}

.taiga-flow-step__body {
  color: var(--taiga-gray-600);
  font-size: 0.9rem;
  margin-bottom: 0.7rem;
  line-height: 1.42;
}

.taiga-tight-block {
  margin-top: 0.55rem;
}

@media (max-width: 900px) {
  .taiga-brandbar {
    grid-template-columns: 82px minmax(0, 1fr);
    gap: 0.8rem;
    align-items: start;
  }

  .taiga-brandbar__logo {
    min-height: 68px;
  }

  .taiga-brandbar__logo img {
    max-height: 52px;
  }

  .taiga-title {
    font-size: 1.7rem;
  }

  .taiga-card {
    min-height: 136px;
  }

  .taiga-card__value {
    min-height: 2.6rem;
  }

  .taiga-card__note {
    min-height: 2.2rem;
  }

  .taiga-subtitle {
    font-size: 0.95rem;
  }

  .element-container:has(.taiga-hero) + div[data-testid="stHorizontalBlock"] {
    display: none;
  }
}

.taiga-line-items__total {
  margin-top: 0.7rem;
  padding-top: 0.7rem;
  border-top: 1px solid rgba(13, 13, 13, 0.12);
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto auto;
  gap: 0.75rem;
  align-items: baseline;
}

.taiga-report {
  background: transparent;
  border: 0;
  border-radius: var(--taiga-radius-md);
  padding: 0.15rem 0;
  box-shadow: none;
}

.taiga-report table {
  width: 100%;
  border-collapse: collapse;
  background: transparent !important;
}

.taiga-report th {
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--taiga-gray-600);
  font-weight: 600;
  padding: 10px 0 12px;
  border-bottom: 2px solid var(--taiga-black);
  text-align: left;
  background: transparent !important;
}

.taiga-report td {
  padding: 12px 0;
  border-bottom: 0;
  font-size: 15px;
  color: var(--taiga-black);
  background: transparent !important;
  font-family: var(--taiga-font-body);
}

.taiga-report td:last-child,
.taiga-report th:last-child {
  text-align: right;
  font-family: var(--taiga-font-body);
}

.taiga-report tr:last-child td {
  padding-bottom: 0;
}

.taiga-report td:first-child {
  letter-spacing: 0.03em;
  text-transform: none;
}

.taiga-config {
  margin-top: 1rem;
  background: linear-gradient(180deg, rgba(232, 242, 236, 0.55), rgba(250, 250, 248, 0.95));
  border: var(--taiga-border);
  border-radius: var(--taiga-radius-md);
  padding: 1rem;
}

.taiga-config__title {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--taiga-gray-400);
  margin-bottom: 0.65rem;
}

.taiga-config__row {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 0.75rem;
  padding: 0.4rem 0;
}

.taiga-config__label {
  color: var(--taiga-gray-600);
  font-size: 0.92rem;
}

.taiga-config__value {
  font-family: var(--taiga-font-mono);
  color: var(--taiga-black);
  font-size: 0.95rem;
  white-space: nowrap;
}

div[data-testid="stNumberInput"],
div[data-testid="stTextInput"],
div[data-testid="stSelectbox"],
div[data-testid="stMultiSelect"],
div[data-testid="stSlider"] {
  max-width: 520px;
}

.taiga-form-section div[data-testid="stNumberInput"],
.taiga-form-section div[data-testid="stTextInput"],
.taiga-form-section div[data-testid="stSelectbox"],
.taiga-form-section div[data-testid="stMultiSelect"],
.taiga-form-section div[data-testid="stSlider"] {
  max-width: 460px;
}

.taiga-form-section div[data-testid="stNumberInput"] > div,
.taiga-form-section div[data-testid="stTextInput"] > div,
.taiga-form-section div[data-testid="stSelectbox"] > div,
.taiga-form-section div[data-testid="stMultiSelect"] > div,
.taiga-form-section div[data-testid="stSlider"] > div {
  max-width: 460px;
}

div[data-testid="stButton"] > button {
  border-radius: 999px;
  border: 1px solid var(--taiga-gray-200);
  background: var(--taiga-white);
  color: var(--taiga-black);
  min-height: 44px;
  font-weight: 500;
}

div[data-testid="stButton"] > button[kind="primary"] {
  background: var(--taiga-green);
  color: var(--taiga-white);
  border-color: var(--taiga-green);
}
"""

WIZARD_STEPS = [
    ("Project Basics", "Define the shared project assumptions."),
    ("Traditional Model", "Set the conventional building baseline."),
    ("Taiga Forma", "Configure products and Taiga operating inputs."),
    ("Taiga Cycle", "Adjust cycle timing and buyback values."),
    ("Leasing", "Review financing assumptions and monthly factors."),
    ("Summary", "Compare outcomes and inspect scenario charts."),
    ("Lifecycle & Reporting", "Review detailed breakdowns and export reports."),
]

CAT_B_REGIONS = {
    "Finland / Norway (Helsinki, Oslo)": {
        "Low": {"min": 1050, "mid": 1200, "max": 1350},
        "Medium": {"min": 2000, "mid": 2450, "max": 2700},
        "High": {"min": 3500, "mid": 4000, "max": 4500},
    },
    "Sweden / Denmark (Stockholm, Copenhagen)": {
        "Low": {"min": 1000, "mid": 1150, "max": 1300},
        "Medium": {"min": 1900, "mid": 2200, "max": 2600},
        "High": {"min": 3200, "mid": 3700, "max": 4200},
    },
    "UK (London)": {
        "Low": {"min": 1500, "mid": 1750, "max": 1950},
        "Medium": {"min": 2700, "mid": 3000, "max": 3300},
        "High": {"min": 4300, "mid": 4900, "max": 5500},
    },
    "Central Europe (Amsterdam, Frankfurt, Paris, Brussels)": {
        "Low": {"min": 800, "mid": 1000, "max": 1200},
        "Medium": {"min": 1500, "mid": 1850, "max": 2200},
        "High": {"min": 2600, "mid": 3200, "max": 3800},
    },
    "Switzerland (Geneva, Zurich)": {
        "Low": {"min": 1300, "mid": 1550, "max": 1800},
        "Medium": {"min": 2400, "mid": 2800, "max": 3200},
        "High": {"min": 3800, "mid": 4400, "max": 5000},
    },
    "Southern Europe (Madrid, Milan, Rome)": {
        "Low": {"min": 550, "mid": 700, "max": 900},
        "Medium": {"min": 1000, "mid": 1300, "max": 1600},
        "High": {"min": 1800, "mid": 2300, "max": 2800},
    },
}

CAT_C_REGIONS = {
    "Finland / Norway (Helsinki, Oslo)": {"min": 24, "max": 69},
    "Sweden / Denmark (Stockholm, Copenhagen)": {"min": 167, "max": 279},
    "UK (London)": {"min": 198, "max": 397},
    "Central Europe (Amsterdam, Frankfurt, Paris, Brussels)": {"min": 108, "max": 233},
    "Switzerland (Geneva, Zurich)": {"min": 237, "max": 296},
    "Southern Europe (Madrid, Milan, Rome)": {"min": 50, "max": 130},
}

LEVEL_OPTIONS = ["Low", "Medium", "High"]

TRAD_FITOUT_ELEMENTS = [
    {"category": "Partitions", "name": "Gypsum board partitions", "unit": "eur_per_jm", "low": (150, 200), "med": (200, 300), "high": (300, 450), "default_qty": "perimeter"},
    {"category": "Partitions", "name": "Framed glass walls", "unit": "eur_per_jm", "low": (0, 0), "med": (400, 700), "high": (600, 900), "default_qty": "perimeter"},
    {"category": "HVAC", "name": "Duct branches and connectors", "unit": "eur_per_room", "low": (0, 0), "med": (800, 1500), "high": (1200, 2000), "default_qty": "rooms"},
    {"category": "HVAC", "name": "VAV controls and thermostat", "unit": "eur_per_room", "low": (0, 0), "med": (600, 1200), "high": (1000, 1800), "default_qty": "rooms"},
    {"category": "HVAC", "name": "System balancing", "unit": "eur_per_room", "low": (0, 0), "med": (300, 600), "high": (500, 900), "default_qty": "rooms"},
    {"category": "Acoustics", "name": "Office acoustic ceilings", "unit": "eur_per_m2", "low": (35, 55), "med": (60, 95), "high": (100, 180), "default_qty": "room_area"},
    {"category": "Lighting", "name": "Office lighting and controls", "unit": "eur_per_m2", "low": (30, 50), "med": (60, 100), "high": (120, 200), "default_qty": "room_area"},
    {"category": "Electrical & Data", "name": "Power outlets and cabling", "unit": "eur_per_room", "low": (300, 500), "med": (400, 800), "high": (700, 1200), "default_qty": "rooms"},
    {"category": "Electrical & Data", "name": "Cat6 data cabling", "unit": "eur_per_room", "low": (200, 400), "med": (300, 600), "high": (500, 900), "default_qty": "rooms"},
    {"category": "Electrical & Data", "name": "Fire detector relocation", "unit": "eur_per_room", "low": (150, 300), "med": (200, 400), "high": (300, 600), "default_qty": "rooms"},
    {"category": "Doors", "name": "Office doors and frames", "unit": "eur_each", "low": (400, 700), "med": (800, 1500), "high": (1500, 3000), "default_qty": "rooms"},
]

TRAD_FEE_ELEMENTS = [
    {"name": "Architect design", "low": (4, 4), "med": (5, 6), "high": (6, 8)},
    {"name": "MEP engineering", "low": (3, 4), "med": (4, 5), "high": (5, 7)},
    {"name": "Project management", "low": (3, 4), "med": (4, 6), "high": (6, 8)},
]


def _format_money(value: float, digits: int = 0) -> str:
    return f"EUR {value:,.{digits}f}"


def _range_for_level(element: dict, level: str) -> tuple[float, float]:
    key = {"Low": "low", "Medium": "med", "High": "high"}[level]
    return element[key]


def _default_fitout_qty(unit_mode: str, room_qty: int, avg_room_m2: float, room_perimeter_jm: float) -> float:
    if unit_mode == "perimeter":
        return float(room_qty) * float(room_perimeter_jm)
    if unit_mode in {"rooms", "rooms_each"}:
        return float(room_qty)
    if unit_mode == "room_area":
        return float(room_qty) * float(avg_room_m2)
    return float(room_qty)


def _compute_trad_fitout_outputs(area_basis: float, use_detailed: bool = False) -> dict:
    ss = st.session_state
    room_qty = max(int(ss.trad_room_qty), 1)
    builder_area = float(max(area_basis, 10.0))
    region = ss.get("trad_cat_region", "Finland / Norway (Helsinki, Oslo)")
    quality = ss.get("trad_cat_quality", "Medium")
    avg_room_m2 = float(ss.get("trad_cat_room_size", 20.0))
    use_derived_perimeter = bool(ss.get("trad_cat_use_derived_perimeter", True))
    derived_perimeter_jm = 4.0 * (avg_room_m2 ** 0.5)
    room_perimeter_jm = derived_perimeter_jm if use_derived_perimeter else float(ss.get("trad_cat_room_perimeter", 18.0))

    detail_rows = []
    subtotal = 0.0
    for idx, element in enumerate(TRAD_FITOUT_ELEMENTS):
        lo, hi = _range_for_level(element, quality)
        default_qty = _default_fitout_qty(element["default_qty"], room_qty, avg_room_m2, room_perimeter_jm)
        qty = float(ss.get(f"trad_fitout_qty_{idx}", default_qty)) if use_detailed else float(default_qty)
        unit_mid = (float(lo) + float(hi)) / 2.0 if (lo or hi) else 0.0
        total_mid = qty * unit_mid
        subtotal += total_mid
        detail_rows.append({
            "Category": element["category"],
            "Element": element["name"],
            "Qty": qty,
            "Unit mid": unit_mid,
            "Cost mid": total_mid,
        })

    fee_total = 0.0
    fee_rows = []
    for fee in TRAD_FEE_ELEMENTS:
        lo, hi = _range_for_level(fee, quality)
        pct_mid = (float(lo) + float(hi)) / 2.0
        fee_cost = subtotal * (pct_mid / 100.0)
        fee_total += fee_cost
        fee_rows.append({"Category": "Fees", "Element": fee["name"], "Qty": pct_mid, "Unit mid": pct_mid, "Cost mid": fee_cost})

    total_cat_b = subtotal + fee_total
    per_m2 = total_cat_b / max(builder_area, 1.0)
    cat_c_mid = (float(CAT_C_REGIONS[region]["min"]) + float(CAT_C_REGIONS[region]["max"])) / 2.0
    eol_pct = cat_c_mid / per_m2 if per_m2 > 0 else 0.0
    cbre_mid = float(CAT_B_REGIONS[region][quality]["mid"])
    return {
        "room_qty": room_qty,
        "builder_area": builder_area,
        "region": region,
        "quality": quality,
        "avg_room_m2": avg_room_m2,
        "use_derived_perimeter": use_derived_perimeter,
        "derived_perimeter_jm": derived_perimeter_jm,
        "room_perimeter_jm": room_perimeter_jm,
        "cbre_mid": cbre_mid,
        "cat_c_mid": cat_c_mid,
        "per_m2": per_m2,
        "eol_pct": eol_pct,
        "total_cat_b": total_cat_b,
        "detail_rows": detail_rows,
        "fee_rows": fee_rows,
    }


def _compute_trad_operational_defaults(area_basis: float) -> dict:
    fitout = _compute_trad_fitout_outputs(area_basis, use_detailed=bool(st.session_state.get("trad_use_fitout_builder", False)))
    base_benchmark = float(CAT_B_REGIONS["Finland / Norway (Helsinki, Oslo)"]["Medium"]["mid"])
    benchmark_factor = fitout["cbre_mid"] / max(base_benchmark, 1.0)
    room_area_factor = fitout["avg_room_m2"] / 20.0
    perimeter_factor = fitout["room_perimeter_jm"] / max(4.0 * (20.0 ** 0.5), 1.0)

    commissioning_cost_unit = 950.0 * benchmark_factor * perimeter_factor
    maint_total_unit = 300.0 * benchmark_factor * (room_area_factor ** 0.5)
    downtime_rate = 15.0 * benchmark_factor
    install_downtime_unit = 80.0 * perimeter_factor
    maint_downtime_unit = 5.0 * (room_area_factor ** 0.5)

    return {
        "commissioning_cost_unit": round(commissioning_cost_unit / 10.0) * 10.0,
        "maint_total_unit": round(maint_total_unit / 10.0) * 10.0,
        "dt_rate": round(downtime_rate, 1),
        "dt_install_h_unit": round(install_downtime_unit, 1),
        "dt_maint_h_total_unit": round(maint_downtime_unit, 1),
        "benchmark_factor": benchmark_factor,
        "room_area_factor": room_area_factor,
        "perimeter_factor": perimeter_factor,
    }


def _render_trad_fitout_builder(area_basis: float):
    st.markdown(
        '<div class="taiga-subsection taiga-subsection--compact">'
        '<div class="taiga-subsection__title">Detailed CAT B / CAT C builder</div>'
        '<div class="taiga-subsection__body">Use a guided benchmark flow to derive a Traditional investment price per m2 and end-of-life share, then fine-tune only if needed.</div>',
        unsafe_allow_html=True,
    )
    st.checkbox("Use builder result for Traditional investment and end-of-life", key="trad_use_fitout_builder")

    room_qty = max(int(st.session_state.trad_room_qty), 1)
    builder_area = float(max(area_basis, 10.0))
    derived_room_m2_raw = builder_area / max(float(room_qty), 1.0)
    derived_room_m2 = max(1.0, derived_room_m2_raw)
    derived_perimeter_jm = 4.0 * (derived_room_m2 ** 0.5)

    if st.session_state.get("trad_cat_auto_geometry", True):
        st.session_state.trad_cat_room_size = float(round(derived_room_m2, 2))
        st.session_state.trad_cat_use_derived_perimeter = True
        st.session_state.trad_cat_room_perimeter = float(round(derived_perimeter_jm, 2))

    st.markdown(
        '<div class="taiga-flow-step taiga-flow-step--compact">'
        '<div class="taiga-flow-step__eyebrow">Step 1</div>'
        '<div class="taiga-flow-step__title">Benchmark</div>'
        '<div class="taiga-flow-step__body">Pick the market and fit-out quality. The builder uses these as the commercial reference point.</div>',
        unsafe_allow_html=True,
    )
    b1, b2 = st.columns(2, gap="small")
    with b1:
        region = st.selectbox("Market region", list(CAT_B_REGIONS.keys()), key="trad_cat_region")
    with b2:
        quality = st.selectbox("Fit-out quality", LEVEL_OPTIONS, index=1, key="trad_cat_quality")

    st.caption(f"Project area is read from Project Basics: {builder_area:,.0f} m2. Closed rooms are read from the Traditional Model input: {room_qty}.")

    benchmark_cols = st.columns(3, gap="small")
    with benchmark_cols[0]:
        _card("Project area basis", f"{builder_area:,.0f} m2", "Read from Project Basics")
    with benchmark_cols[1]:
        _card("Closed rooms basis", f"{room_qty:,}", "Read from Traditional Model")
    with benchmark_cols[2]:
        _card("Market benchmark", f"EUR {float(CAT_B_REGIONS[region][quality]['mid']):,.0f} / m2", f"{quality} in {region}")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="taiga-flow-step taiga-flow-step--compact">'
        '<div class="taiga-flow-step__eyebrow">Step 2</div>'
        '<div class="taiga-flow-step__title">Automatic estimate</div>'
        '<div class="taiga-flow-step__body">Start with automatically derived room geometry from project area and room count. Switch to manual only when you need to override the assumption.</div>',
        unsafe_allow_html=True,
    )
    st.checkbox(
        "Auto-calculate room geometry from area and room count",
        key="trad_cat_auto_geometry",
        help="Keeps room size and perimeter aligned with the main project area and Traditional room count.",
    )

    derived_top = st.columns(2, gap="small")
    with derived_top[0]:
        _card("Derived room size", f"{derived_room_m2_raw:,.2f} m2", "Project area / closed rooms")
    with derived_top[1]:
        _card("Derived room perimeter", f"{derived_perimeter_jm:,.2f} jm", "Square-room approximation")

    warning_parts = []
    if derived_room_m2_raw < 5.0:
        warning_parts.append("Derived room size is very small for a closed room.")
    if derived_room_m2_raw > 100.0:
        warning_parts.append("Derived room size is very large for a single closed room.")
    if derived_perimeter_jm > 60.0:
        warning_parts.append("Derived room perimeter is unusually large.")
    if warning_parts:
        st.markdown(
            f'<div class="taiga-inline-note taiga-inline-note--warning">Geometry warning: {" ".join(warning_parts)} Check project area and room count, or switch to manual geometry.</div>',
            unsafe_allow_html=True,
        )

    geometry_cols = st.columns(2, gap="small")
    with geometry_cols[0]:
        st.number_input(
            "Average room size (m2)",
            min_value=1.0,
            max_value=1000.0,
            value=float(min(max(st.session_state.get("trad_cat_room_size", 20.0), 1.0), 1000.0)),
            step=1.0,
            key="trad_cat_room_size",
            disabled=bool(st.session_state.get("trad_cat_auto_geometry", True)),
            help="When auto-calculate is enabled, this value is derived from project area divided by closed rooms.",
        )
    with geometry_cols[1]:
        st.checkbox(
            "Derive room perimeter from room size",
            key="trad_cat_use_derived_perimeter",
            help="Uses a square-room approximation: perimeter = 4 x sqrt(room area).",
            disabled=bool(st.session_state.get("trad_cat_auto_geometry", True)),
        )

    perimeter_cols = st.columns(2, gap="small")
    with perimeter_cols[0]:
        st.number_input(
            "Average room perimeter (jm)",
            min_value=1.0,
            max_value=250.0,
            value=float(min(max(st.session_state.get("trad_cat_room_perimeter", 18.0), 1.0), 250.0)),
            step=1.0,
            key="trad_cat_room_perimeter",
            disabled=bool(st.session_state.get("trad_cat_auto_geometry", True) or st.session_state.get("trad_cat_use_derived_perimeter", True)),
            help="Manual override for perimeter-sensitive elements such as partitions and framed glass walls.",
        )
    with perimeter_cols[1]:
        source_mode = "Automatic from project area and room count" if st.session_state.get("trad_cat_auto_geometry", True) else ("Derived from room size" if st.session_state.get("trad_cat_use_derived_perimeter", True) else "Manual geometry")
        st.markdown(
            f'<div class="taiga-inline-note">Geometry mode in use: {source_mode}</div>',
            unsafe_allow_html=True,
        )

    auto_preview = _compute_trad_fitout_outputs(area_basis, use_detailed=False)
    preview_cols = st.columns(2, gap="small")
    with preview_cols[0]:
        source_label = "Derived from room size" if auto_preview["use_derived_perimeter"] else "Manual perimeter"
        _card("Room size used", f"{auto_preview['avg_room_m2']:,.1f} m2", "Current estimate basis")
    with preview_cols[1]:
        _card("Perimeter used", f"{auto_preview['room_perimeter_jm']:,.1f} jm", source_label)

    recommendation_per_m2 = auto_preview["per_m2"]
    recommendation_eol_pct = auto_preview["eol_pct"]
    applied_preview = _compute_trad_fitout_outputs(area_basis, use_detailed=bool(st.session_state.get("trad_use_fitout_builder", False)))

    summary_top = st.columns(2, gap="small")
    with summary_top[0]:
        _card("Recommended investment", f"EUR {recommendation_per_m2:,.0f} / m2", "Auto-estimated from current geometry", dark=True)
    with summary_top[1]:
        _card("Recommended end-of-life", f"{recommendation_eol_pct * 100:,.1f}%", "Share of investment from CAT C midpoint")
    summary_bottom = st.columns(2, gap="small")
    with summary_bottom[0]:
        _card("Benchmark reference", f"EUR {auto_preview['cbre_mid']:,.0f} / m2", f"{quality} in {region}")
    with summary_bottom[1]:
        _card("CAT C midpoint", f"EUR {auto_preview['cat_c_mid']:,.0f} / m2", "Reinstatement midpoint")

    applied_note = "Applied to Traditional Model inputs" if st.session_state.get("trad_use_fitout_builder") else "Preview only until you enable the builder result"
    st.markdown(
        f'<div class="taiga-inline-note">Current builder output: EUR {applied_preview["per_m2"]:,.0f} / m2 and {applied_preview["eol_pct"] * 100:,.1f}% end-of-life. {applied_note}.</div>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="taiga-flow-step taiga-flow-step--compact">'
        '<div class="taiga-flow-step__eyebrow">Step 3</div>'
        '<div class="taiga-flow-step__title">Optional detailed adjustment</div>'
        '<div class="taiga-flow-step__body">Open this only when you need to fine-tune quantities beyond the automatic estimate. The detailed result can still drive the same Traditional investment price per m2.</div>',
        unsafe_allow_html=True,
    )
    with st.expander("Adjust detailed assumptions", expanded=False):
        st.checkbox(
            "Keep detailed assumptions synced with the automatic estimate",
            key="trad_fitout_auto_sync",
            help="When enabled, detailed quantities follow the current automatic room geometry defaults. Turn off only when you want to custom-edit quantities by hand.",
        )
        tuned_rows = []
        tuned_subtotal = 0.0
        categories = list(dict.fromkeys(element["category"] for element in TRAD_FITOUT_ELEMENTS))
        for category in categories:
            with st.expander(category, expanded=(category == categories[0])):
                for idx, element in enumerate(TRAD_FITOUT_ELEMENTS):
                    if element["category"] != category:
                        continue
                    lo, hi = _range_for_level(element, quality)
                    default_qty = _default_fitout_qty(
                        element["default_qty"],
                        int(room_qty),
                        float(auto_preview["avg_room_m2"]),
                        float(auto_preview["room_perimeter_jm"]),
                    )
                    if st.session_state.get("trad_fitout_auto_sync", True):
                        st.session_state[f"trad_fitout_qty_{idx}"] = float(default_qty)
                    qty_cols = st.columns([0.62, 0.18, 0.20], gap="small")
                    with qty_cols[0]:
                        st.markdown(f"**{element['name']}**")
                        basis_note = {
                            "perimeter": "Calculated from closed rooms x room perimeter",
                            "rooms": "Calculated from closed room count",
                            "room_area": "Calculated from enclosed room area",
                            "rooms_each": "Calculated from closed room count",
                        }.get(element["default_qty"], "Calculated from current builder assumptions")
                        st.caption(f"Default quantity: {default_qty:,.0f} | {basis_note}")
                    with qty_cols[1]:
                        qty = st.number_input(
                            f"{element['name']} quantity",
                            min_value=0.0,
                            max_value=100000.0,
                            value=float(default_qty),
                            step=1.0,
                            key=f"trad_fitout_qty_{idx}",
                            label_visibility="collapsed",
                            help="Override the automatically derived quantity for this element.",
                        )
                    with qty_cols[2]:
                        unit_mid = (float(lo) + float(hi)) / 2.0 if (lo or hi) else 0.0
                        total_mid = qty * unit_mid
                        st.markdown(f"**EUR {total_mid:,.0f}**")
                    tuned_subtotal += total_mid
                    tuned_rows.append({
                        "Category": element["category"],
                        "Element": element["name"],
                        "Qty": qty,
                        "Unit mid": unit_mid,
                        "Cost mid": total_mid,
                    })

        tuned_fee_total = 0.0
        tuned_fee_rows = []
        for fee in TRAD_FEE_ELEMENTS:
            lo, hi = _range_for_level(fee, quality)
            pct_mid = (float(lo) + float(hi)) / 2.0
            fee_cost = tuned_subtotal * (pct_mid / 100.0)
            tuned_fee_total += fee_cost
            tuned_fee_rows.append({"Category": "Fees", "Element": fee["name"], "Qty": pct_mid, "Unit mid": pct_mid, "Cost mid": fee_cost})

        tuned_total_cat_b = tuned_subtotal + tuned_fee_total
        tuned_per_m2 = tuned_total_cat_b / max(float(builder_area), 1.0)
        tuned_eol_pct = auto_preview["cat_c_mid"] / tuned_per_m2 if tuned_per_m2 > 0 else 0.0

        fine_cols = st.columns(2, gap="small")
        with fine_cols[0]:
            _card("Detailed investment (EUR / m2)", f"EUR {tuned_per_m2:,.0f}", "Based on adjusted quantities")
        with fine_cols[1]:
            _card("Detailed end-of-life", f"{tuned_eol_pct * 100:,.1f}%", "Updated from adjusted investment")

        if st.session_state.get("trad_use_fitout_builder"):
            st.markdown(
                f'<div class="taiga-inline-note">Detailed assumptions are active. Traditional investment price per m2 now uses EUR {tuned_per_m2:,.0f} / m2.</div>',
                unsafe_allow_html=True,
            )

        detail_df = pd.DataFrame(tuned_rows + tuned_fee_rows)
        if not detail_df.empty:
            render_df = detail_df.copy()
            render_df["Qty"] = render_df["Qty"].map(lambda x: f"{x:,.0f}")
            render_df["Unit mid"] = render_df["Unit mid"].map(lambda x: f"{x:,.0f}")
            render_df["Cost mid"] = render_df["Cost mid"].map(lambda x: f"{x:,.0f}")
            st.markdown("#### Builder cost summary")
            st.markdown(f'<div class="taiga-report">{render_df.to_html(index=False, classes="taiga-table", border=0)}</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


def _get_effective_area(values: dict) -> float:
    area_from_products = values.get("area_from_products")
    if values.get("override_area") and area_from_products is not None:
        return float(area_from_products)
    return float(values.get("area_m2", 0.0))


def _snapshot_state() -> dict:
    ss = st.session_state
    return {
        "years": int(ss.years),
        "wacc": float(ss.wacc),
        "area_m2": float(ss.area_m2),
        "kwh_m2yr": float(ss.kwh_m2yr),
        "elec_price": float(ss.elec_price),
        "cycle_year": int(ss.cycle_year),
        "area_from_products": ss.get("area_from_products"),
        "override_area": bool(ss.get("override_area", False)),
        "override_price": bool(ss.get("override_price", False)),
        "taiga_list_price": float(ss.taiga_list_price),
        "taiga_occ_rate": float(ss.taiga_occ_rate),
        "taiga_standby": float(ss.taiga_standby),
        "taiga_commissioning_cost_unit": float(ss.taiga_commissioning_cost_unit),
        "taiga_maint_total_unit": float(ss.taiga_maint_total_unit),
        "taiga_dt_rate": float(ss.taiga_dt_rate),
        "taiga_dt_install_h_unit": float(ss.taiga_dt_install_h_unit),
        "taiga_dt_maint_h_total_unit": float(ss.taiga_dt_maint_h_total_unit),
        "taiga_commissioning_year": int(ss.taiga_commissioning_year),
        "taiga_eol_cost": float(ss.taiga_eol_cost),
        "taiga_total_qty": int(ss.get("taiga_total_qty", 0)),
        "trad_price_per_m2": float(ss.trad_price_per_m2),
        "trad_commissioning_cost_unit": float(ss.trad_commissioning_cost_unit),
        "trad_commissioning_year": int(ss.trad_commissioning_year),
        "trad_maint_total_unit": float(ss.trad_maint_total_unit),
        "trad_dt_rate": float(ss.trad_dt_rate),
        "trad_dt_install_h_unit": float(ss.trad_dt_install_h_unit),
        "trad_dt_maint_h_total_unit": float(ss.trad_dt_maint_h_total_unit),
        "trad_eol_pct": float(ss.trad_eol_pct),
        "trad_run_frac": float(ss.trad_run_frac),
        "trad_room_qty": int(ss.trad_room_qty),
        "lease_term_years": int(ss.lease_term_years),
        "lease_wacc_annual": float(ss.lease_wacc_annual),
        "lease_base_price": float(ss.lease_base_price),
        "lease_buyback_year": int(ss.lease_buyback_year),
        "cycle_df": ss.cycle_df.copy(),
        "price_df": ss.price_df.copy(),
        "lease_factors_df": ss.lease_factors_df.copy(),
    }


def _sync_ui_state():
    ss = st.session_state
    ss.setdefault("wacc_pct_ui", float(ss.wacc) * 100.0)
    ss.setdefault("lease_wacc_pct_ui", float(ss.lease_wacc_annual) * 100.0)


def _render_percent_input(label: str, source_key: str, ui_key: str, max_value: float = 100.0, step: float = 0.1):
    value = st.number_input(label, min_value=0.0, max_value=max_value, step=step, key=ui_key, format="%.2f")
    st.session_state[source_key] = float(value) / 100.0
    return value


def _render_report_table(df: pd.DataFrame, title: str, digits: int = 0, index: bool = True):
    st.markdown(f"#### {title}")
    if df is None or df.empty:
        st.info("No data available for this view.")
        return
    render_df = df.copy()
    if index:
        render_df.index = [str(idx).replace("_", " ").title() for idx in render_df.index]
    render_df.columns = [str(col).replace("_", " ").title() for col in render_df.columns]
    for col in render_df.columns:
        if pd.api.types.is_numeric_dtype(render_df[col]):
            render_df[col] = render_df[col].map(lambda x: f"{x:,.{digits}f}")
    table_html = render_df.to_html(index=index, classes="taiga-table", border=0)
    st.markdown(f'<div class="taiga-report">{table_html}</div>', unsafe_allow_html=True)


def _load_price_list_upload(key: str = "taiga_price_upload", label: str = "Upload price list (xlsx / csv)"):
    upload = st.file_uploader(label, type=["xlsx", "csv"], key=key)
    if upload is None:
        return
    try:
        if upload.name.lower().endswith(".xlsx"):
            df = pd.read_excel(upload)
        else:
            df = pd.read_csv(upload)
        cols_map = {c.lower(): c for c in df.columns}
        required = {"code", "name", "unit_price_eur"}
        if not required.issubset(set(cols_map.keys())):
            st.error("File must contain: code, name, unit_price_eur. area_m2 is optional.")
            return
        order = ["code", "name", "unit_price_eur"] + (["area_m2"] if "area_m2" in cols_map else [])
        df = df[[cols_map[c] for c in order]].copy()
        if "qty" not in df.columns:
            df["qty"] = 0
        st.session_state.price_df = df.copy()
    except Exception as exc:
        st.error(f"Failed to read file: {exc}")


def _product_family(code: str) -> str:
    code = str(code).upper()
    if code.startswith("LB"):
        return "LohkoBox"
    if code.startswith("FL"):
        return "Flex"
    if code.startswith("PIC"):
        return "Picea"
    return "Other"


def _render_taiga_product_selector():
    price_df = st.session_state.price_df.copy()
    if price_df.empty:
        st.info("No products available.")
        return

    if "area_m2" not in price_df.columns:
        price_df["area_m2"] = 0.0
    if "qty" not in price_df.columns:
        price_df["qty"] = 0
    price_df["family"] = price_df["code"].map(_product_family)
    price_df["name"] = price_df["name"].astype(str)
    price_df["code"] = price_df["code"].astype(str)

    st.markdown(
        """
        <div class="taiga-picker">
          <div class="taiga-picker__title">Product workflow</div>
          <div class="taiga-picker__body">Filter the catalog, choose products, then adjust only the quantities you want to include in the Taiga concept.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    filter_cols = st.columns([0.34, 0.66], gap="medium")
    with filter_cols[0]:
        family_options = ["All"] + sorted(price_df["family"].dropna().unique().tolist())
        selected_family = st.selectbox("Product family", family_options, key="taiga_family_filter")
    with filter_cols[1]:
        search_term = st.text_input("Search products", key="taiga_product_search", placeholder="Search by code or product name")

    filtered_df = price_df.copy()
    if selected_family != "All":
        filtered_df = filtered_df.loc[filtered_df["family"] == selected_family].copy()
    if search_term.strip():
        needle = search_term.strip().lower()
        filtered_df = filtered_df.loc[
            filtered_df["code"].str.lower().str.contains(needle) |
            filtered_df["name"].str.lower().str.contains(needle)
        ].copy()

    option_map = {
        row.code: f"{row.code} â€¢ {row.name} â€¢ EUR {float(row.unit_price_eur):,.0f} / unit â€¢ {float(row.area_m2):,.0f} m2"
        for row in filtered_df.itertuples(index=False)
    }
    all_selected_codes = price_df.loc[pd.to_numeric(price_df["qty"], errors="coerce").fillna(0) > 0, "code"].astype(str).tolist()
    visible_codes = list(option_map.keys())
    default_codes = [code for code in all_selected_codes if code in visible_codes]
    visible_selected_codes = st.multiselect(
        "Choose Taiga Forma products",
        options=visible_codes,
        default=default_codes,
        format_func=lambda code: option_map.get(code, code),
        key="taiga_selected_codes",
    )
    hidden_selected_codes = [code for code in all_selected_codes if code not in visible_codes]
    selected_codes = list(dict.fromkeys(visible_selected_codes + hidden_selected_codes))

    original_df = price_df.copy()
    updated_df = price_df.copy()
    updated_df["qty"] = 0
    if selected_codes:
        pill_html = []
        for code in selected_codes:
            row = original_df.loc[original_df["code"] == code].iloc[0]
            default_qty = int(row["qty"]) if int(row["qty"]) > 0 else 1
            family = _product_family(code)
            pill_html.append(
                f'<div class="taiga-selected-pill">'
                f'<span class="taiga-family-badge">{family}</span>{code}'
                f'<span class="taiga-selected-pill__qty">Qty {default_qty}</span>'
                f"</div>"
            )
        st.markdown("##### Selected Taiga Forma")
        st.markdown(f'<div class="taiga-selected-pills">{"".join(pill_html)}</div>', unsafe_allow_html=True)
        st.markdown("##### Quantities")
        for code in selected_codes:
            row = original_df.loc[original_df["code"] == code].iloc[0]
            family = _product_family(code)
            qty_cols = st.columns([0.60, 0.18, 0.22], gap="small")
            with qty_cols[0]:
                st.markdown(
                    f'<div class="taiga-qty-row__title"><span class="taiga-family-badge">{family}</span>{row["code"]} - {row["name"]}</div>'
                    f'<div class="taiga-qty-row__meta">Area {float(row["area_m2"]):,.0f} m2 per unit | EUR {float(row["unit_price_eur"]):,.0f} / unit</div>',
                    unsafe_allow_html=True,
                )
            with qty_cols[1]:
                qty_value = st.number_input(
                    f"{code} quantity",
                    min_value=0,
                    max_value=1000,
                    value=int(row["qty"]) if int(row["qty"]) > 0 else 1,
                    step=1,
                    key=f"taiga_qty_{code}",
                )
            with qty_cols[2]:
                st.markdown(
                    f'<div class="taiga-qty-row__value">EUR {float(qty_value) * float(row["unit_price_eur"]):,.0f}</div>',
                    unsafe_allow_html=True,
                )
            updated_df.loc[updated_df["code"] == code, "qty"] = int(qty_value)

    if not filtered_df.empty:
        st.markdown('<div class="taiga-catalog">', unsafe_allow_html=True)
        for family in filtered_df["family"].dropna().unique().tolist():
            family_df = filtered_df.loc[filtered_df["family"] == family].copy()
            cards = [f'<div class="taiga-family-section"><div class="taiga-family-title">{family}</div>']
            for row in family_df.itertuples(index=False):
                qty_series = pd.to_numeric(price_df.loc[price_df["code"] == row.code, "qty"], errors="coerce").fillna(0)
                qty_value = int(qty_series.iloc[0]) if not qty_series.empty else 0
                selected_marker = ('<span class="taiga-catalog-card__status"><span class="taiga-catalog-card__status-pill taiga-catalog-card__status-pill--active">Included</span>' + f'<span class="taiga-catalog-card__status-pill taiga-catalog-card__status-pill--active">Qty {qty_value}</span></span>') if row.code in selected_codes and qty_value > 0 else '<span class="taiga-catalog-card__status"><span class="taiga-catalog-card__status-pill">Available</span></span>'
                cards.append(
                    f'<div class="taiga-catalog-card">'
                    f'<div class="taiga-catalog-card__head">'
                    f'<div class="taiga-catalog-card__title"><span class="taiga-family-badge">{family}</span>{row.code} - {row.name}</div>'
                    f'<div class="taiga-catalog-card__price">EUR {float(row.unit_price_eur):,.0f}</div>'
                    f'</div>'
                    f'<div class="taiga-catalog-card__meta">{float(row.area_m2):,.0f} m2 per unit {selected_marker}</div>'
                    f'</div>'
                )
            cards.append("</div>")
            st.markdown("".join(cards), unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    if not selected_codes:
        st.info("Choose one or more products to build the Taiga concept package.")

    non_selected_codes = set(updated_df["code"]) - set(selected_codes)
    if non_selected_codes:
        updated_df.loc[updated_df["code"].isin(non_selected_codes), "qty"] = 0

    st.session_state.price_df = updated_df.copy()

    selected_df = updated_df.loc[updated_df["qty"] > 0].copy()
    total_qty = int(selected_df["qty"].sum()) if not selected_df.empty else 0
    total_price = float((selected_df["qty"] * selected_df["unit_price_eur"]).sum()) if not selected_df.empty else 0.0
    total_area = float((selected_df["qty"] * selected_df["area_m2"]).sum()) if not selected_df.empty else 0.0

    st.session_state.taiga_total_qty = total_qty
    st.session_state.area_from_products = total_area if total_qty > 0 else None
    st.session_state.override_price = True
    st.session_state.taiga_list_price = total_price

    st.session_state["taiga_selected_units"] = total_qty
    st.session_state["taiga_selected_area"] = total_area
    st.session_state["taiga_selected_list_price"] = total_price



def _render_selected_products_summary():
    selected_df = st.session_state.price_df.copy()
    if selected_df.empty:
        return
    selected_df["qty"] = pd.to_numeric(selected_df["qty"], errors="coerce").fillna(0)
    selected_df["unit_price_eur"] = pd.to_numeric(selected_df["unit_price_eur"], errors="coerce").fillna(0.0)
    if "area_m2" not in selected_df.columns:
        selected_df["area_m2"] = 0.0
    selected_df["area_m2"] = pd.to_numeric(selected_df["area_m2"], errors="coerce").fillna(0.0)
    selected_df = selected_df.loc[selected_df["qty"] > 0].copy()
    if selected_df.empty:
        return

    rows = []
    total_units = int(selected_df["qty"].sum())
    total_area = float((selected_df["qty"] * selected_df["area_m2"]).sum())
    total_value = float((selected_df["qty"] * selected_df["unit_price_eur"]).sum())
    for row in selected_df.itertuples(index=False):
        line_total = float(row.qty) * float(row.unit_price_eur)
        family = _product_family(row.code)
        rows.append(
            f'<div class="taiga-line-item">'
            f"<div>"
            f'<div class="taiga-line-item__title"><span class="taiga-family-badge">{family}</span>{row.code} - {row.name}</div>'
            f'<div class="taiga-line-item__meta">Area {float(row.area_m2):,.0f} m2 per unit</div>'
            f"</div>"
            f'<div class="taiga-line-item__qty">{int(row.qty)} pcs</div>'
            f'<div class="taiga-line-item__value">EUR {line_total:,.0f}</div>'
            f"</div>"
        )

    st.markdown(
        f'<div class="taiga-line-items">'
        f'<div class="taiga-line-items__header">Selected Taiga Forma package</div>'
        f'{"".join(rows)}'
        f'<div class="taiga-line-items__total">'
        f"<div>"
        f'<div class="taiga-line-item__title">Taiga Forma total</div>'
        f'<div class="taiga-line-item__meta">Combined area {total_area:,.2f} m2</div>'
        f"</div>"
        f'<div class="taiga-line-item__qty">{total_units} pcs</div>'
        f'<div class="taiga-line-item__value">EUR {total_value:,.0f}</div>'
        f"</div>"
        f"</div>",
        unsafe_allow_html=True,
    )


def _render_taiga_configuration_card():
    selected_df = st.session_state.price_df.copy()
    if selected_df.empty:
        selected_df = pd.DataFrame(columns=["qty", "area_m2", "unit_price_eur"])
    selected_df["qty"] = pd.to_numeric(selected_df.get("qty", 0), errors="coerce").fillna(0)
    selected_df["unit_price_eur"] = pd.to_numeric(selected_df.get("unit_price_eur", 0.0), errors="coerce").fillna(0.0)
    if "area_m2" not in selected_df.columns:
        selected_df["area_m2"] = 0.0
    selected_df["area_m2"] = pd.to_numeric(selected_df["area_m2"], errors="coerce").fillna(0.0)
    selected_df = selected_df.loc[selected_df["qty"] > 0].copy()

    total_units = int(selected_df["qty"].sum()) if not selected_df.empty else 0
    total_area = float((selected_df["qty"] * selected_df["area_m2"]).sum()) if not selected_df.empty else 0.0
    total_price = float((selected_df["qty"] * selected_df["unit_price_eur"]).sum()) if not selected_df.empty else 0.0
    effective_area_basis = max(float(_get_effective_area(_snapshot_state())), 1.0)
    effective_area_share = (total_area / effective_area_basis) * 100.0
    override_area = "Selected products" if st.session_state.get("override_area") else "Shared project area"

    st.markdown(
        f"""
        <div class="taiga-config">
          <div class="taiga-config__title">Taiga Forma configuration</div>
          <div class="taiga-config__row">
            <div class="taiga-config__label">Products selected</div>
            <div class="taiga-config__value">{len(selected_df)}</div>
          </div>
          <div class="taiga-config__row">
            <div class="taiga-config__label">Units included</div>
            <div class="taiga-config__value">{total_units}</div>
          </div>
            <div class="taiga-config__row">
              <div class="taiga-config__label">Combined area</div>
              <div class="taiga-config__value">{total_area:,.2f} m2</div>
            </div>
            <div class="taiga-config__row">
              <div class="taiga-config__label">Effective area share</div>
              <div class="taiga-config__value">{effective_area_share:,.1f}% of effective area basis</div>
            </div>
            <div class="taiga-config__row">
              <div class="taiga-config__label">Catalog value</div>
              <div class="taiga-config__value">EUR {total_price:,.0f}</div>
            </div>
          <div class="taiga-config__row">
            <div class="taiga-config__label">Pricing source</div>
            <div class="taiga-config__value">Selected products</div>
          </div>
          <div class="taiga-config__row">
            <div class="taiga-config__label">Area source</div>
            <div class="taiga-config__value">{override_area}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _build_inputs_taiga(values: dict, to_cycle_list):
    years = int(values["years"])
    qty = int(values["taiga_total_qty"])
    return TCOInputs(
        is_taiga=True,
        years=years,
        wacc=float(values["wacc"]),
        list_price=float(values["taiga_list_price"]),
        area_m2=float(_get_effective_area(values)),
        kwh_m2yr=float(values["kwh_m2yr"]),
        elec_price=float(values["elec_price"]),
        run_frac_trad=0.0,
        occ_rate=float(values["taiga_occ_rate"]),
        standby_taiga=float(values["taiga_standby"]),
        commissioning_cost=float(values["taiga_commissioning_cost_unit"]) * qty,
        commissioning_year=int(values["taiga_commissioning_year"]),
        maint_total=float(values["taiga_maint_total_unit"]) * qty * years,
        downtime_rate_per_hour=float(values["taiga_dt_rate"]),
        downtime_hours_install=float(values["taiga_dt_install_h_unit"]) * qty,
        downtime_hours_maint_total=float(values["taiga_dt_maint_h_total_unit"]) * qty * years,
        eol_cost=float(values["taiga_eol_cost"]),
        cycle_table=to_cycle_list(values["cycle_df"]),
        cycle_year=int(values["cycle_year"]),
    )


def _build_inputs_trad(values: dict):
    years = int(values["years"])
    qty = int(values["trad_room_qty"])
    area = float(_get_effective_area(values))
    list_price_total = float(values["trad_price_per_m2"]) * area
    return TCOInputs(
        is_taiga=False,
        years=years,
        wacc=float(values["wacc"]),
        list_price=list_price_total,
        area_m2=area,
        kwh_m2yr=float(values["kwh_m2yr"]),
        elec_price=float(values["elec_price"]),
        run_frac_trad=float(values["trad_run_frac"]),
        occ_rate=0.0,
        standby_taiga=0.0,
        commissioning_cost=float(values["trad_commissioning_cost_unit"]) * qty,
        commissioning_year=int(values["trad_commissioning_year"]),
        maint_total=float(values["trad_maint_total_unit"]) * qty * years,
        downtime_rate_per_hour=float(values["trad_dt_rate"]),
        downtime_hours_install=float(values["trad_dt_install_h_unit"]) * qty,
        downtime_hours_maint_total=float(values["trad_dt_maint_h_total_unit"]) * qty * years,
        eol_cost=float(values["trad_eol_pct"]) * list_price_total,
        cycle_table=[],
        cycle_year=0,
    )


def _ensure_downtime(df: pd.DataFrame, inp, ensure_cost_columns, ensure_year_row) -> pd.DataFrame:
    if df is None or df.empty or inp is None:
        return df
    df = ensure_cost_columns(df)
    years = int(getattr(inp, "years", 0) or 0)
    for year in range(0, years + 1):
        df = ensure_year_row(df, year)
    if "downtime_pv" not in df.columns:
        df["downtime_pv"] = 0.0

    r = float(getattr(inp, "downtime_rate_per_hour", 0.0) or 0.0)
    h_install = float(getattr(inp, "downtime_hours_install", 0.0) or 0.0)
    h_maint = float(getattr(inp, "downtime_hours_maint_total", 0.0) or 0.0)
    wacc = float(getattr(inp, "wacc", 0.0) or 0.0)
    eps = 1e-6

    if r > 0 and h_install > 0:
        cur0 = float(df.loc[df["year"] == 0, "downtime_pv"].sum() or 0.0)
        if abs(cur0) < eps:
            df.loc[df["year"] == 0, "downtime_pv"] = df.loc[df["year"] == 0, "downtime_pv"] + (r * h_install)

    if r > 0 and h_maint > 0:
        for year in range(1, years + 1):
            current = float(df.loc[df["year"] == year, "downtime_pv"].sum() or 0.0)
            if abs(current) < eps:
                df.loc[df["year"] == year, "downtime_pv"] = df.loc[df["year"] == year, "downtime_pv"] + ((r * h_maint) / ((1.0 + wacc) ** year))
    return ensure_cost_columns(df)


def _build_products_for_offer(values: dict) -> pd.DataFrame:
    src = values.get("price_df", pd.DataFrame()).copy()
    if src is None or src.empty:
        return pd.DataFrame(columns=["product", "qty", "unit_price", "discount_pct"])
    src["qty"] = pd.to_numeric(src["qty"], errors="coerce").fillna(0)
    src["unit_price_eur"] = pd.to_numeric(src["unit_price_eur"], errors="coerce").fillna(0.0)
    sel = src[src["qty"] > 0].copy()
    if sel.empty:
        return pd.DataFrame(columns=["product", "qty", "unit_price", "discount_pct"])
    if "name" in sel.columns and "code" in sel.columns:
        product_series = sel["name"].astype(str).where(sel["name"].astype(str).str.strip() != "", sel["code"].astype(str))
    elif "name" in sel.columns:
        product_series = sel["name"].astype(str)
    elif "code" in sel.columns:
        product_series = sel["code"].astype(str)
    else:
        product_series = sel.index.astype(str)
    return pd.DataFrame({
        "product": product_series,
        "qty": sel["qty"].astype(float),
        "unit_price": sel["unit_price_eur"].astype(float),
        "discount_pct": 0.0,
    })


def _compute_leasing(values: dict) -> dict:
    cp_list = []
    for row in values["cycle_df"].itertuples(index=False):
        raw = float(getattr(row, "value_pct"))
        cp_list.append(lc.CyclePoint(year=int(getattr(row, "year")), value_pct=raw / 100.0 if abs(raw) > 1.0 else raw))

    mo_factor_raw = lc.monthly_factor_for_term(int(values["lease_term_years"]), values["lease_factors_df"])
    mo_factor = mo_factor_raw / 100.0 if mo_factor_raw > 1 else mo_factor_raw
    base_mo, mo_with_buyback = lc.monthly_payment_with_buyback(
        list_price=float(values["lease_base_price"]),
        monthly_factor=mo_factor,
        wacc_annual=float(values["lease_wacc_annual"]),
        term_years=int(values["lease_term_years"]),
        cycle_year=int(values["lease_buyback_year"]),
        cycle_table=cp_list,
    )
    yearly = lc.leasing_yearly_pv_table(
        list_price=float(values["lease_base_price"]),
        monthly_factor=mo_factor,
        wacc_annual=float(values["lease_wacc_annual"]),
        term_years=int(values["lease_term_years"]),
        cycle_year=int(values["lease_buyback_year"]),
        cycle_table=cp_list,
    )
    return {
        "monthly_factor": mo_factor,
        "base_monthly": base_mo,
        "monthly_with_buyback": mo_with_buyback,
        "term_months": int(values["lease_term_years"]) * 12,
        "df_yearly": yearly,
        "pivot": lc.pivot_leasing_for_display(yearly).round(0),
    }


def _compute_results(values, ctx) -> dict:
    ensure_cost_columns = ctx["ensure_cost_columns"]
    ensure_year_row = ctx["ensure_year_row"]
    yearly_breakdown = ctx["yearly_breakdown"]
    component_summary_from_yearly = ctx["component_summary_from_yearly"]
    pv_total_from_yearly = ctx["pv_total_from_yearly"]
    pivot_for_display = ctx["pivot_for_display"]
    to_cycle_list = ctx["to_cycle_list"]

    taiga_inp = _build_inputs_taiga(values, to_cycle_list)
    trad_inp = _build_inputs_trad(values)

    df_taiga = pd.DataFrame(yearly_breakdown(taiga_inp))
    df_trad = pd.DataFrame(yearly_breakdown(trad_inp))
    df_taiga = ensure_cost_columns(_ensure_downtime(ensure_year_row(df_taiga, 0), taiga_inp, ensure_cost_columns, ensure_year_row))
    df_trad = ensure_cost_columns(_ensure_downtime(ensure_year_row(df_trad, 0), trad_inp, ensure_cost_columns, ensure_year_row))

    if "acquisition_pv" not in df_taiga.columns:
        df_taiga["acquisition_pv"] = 0.0
    if "acquisition_pv" not in df_trad.columns:
        df_trad["acquisition_pv"] = 0.0
    df_taiga.loc[df_taiga["year"] == 0, "acquisition_pv"] = float(taiga_inp.list_price)
    df_trad.loc[df_trad["year"] == 0, "acquisition_pv"] = float(trad_inp.list_price)

    buyback_year = int(values["cycle_year"])
    if buyback_year > 0:
        row = values["cycle_df"].loc[values["cycle_df"]["year"] == buyback_year]
        if not row.empty:
            raw = float(row["value_pct"].iloc[0])
            pct = raw / 100.0 if abs(raw) > 1.0 else raw
            if pct > 0:
                df_taiga = ensure_year_row(df_taiga, buyback_year)
                if "buyback_pv" not in df_taiga.columns:
                    df_taiga["buyback_pv"] = 0.0
                df_taiga.loc[df_taiga["year"] == buyback_year, "buyback_pv"] = -float(taiga_inp.list_price) * pct / ((1.0 + float(taiga_inp.wacc)) ** buyback_year)

    df_taiga = ensure_cost_columns(df_taiga)
    df_trad = ensure_cost_columns(df_trad)

    df_delta = None
    if not df_taiga.empty and not df_trad.empty:
        common_cols = [c for c in df_taiga.columns if c in df_trad.columns]
        sum_cols = [c for c in common_cols if c != "year"]
        merged = df_taiga[["year"] + sum_cols].merge(df_trad[["year"] + sum_cols], on="year", suffixes=("_T", "_TR"))
        for col in sum_cols:
            merged[col] = merged[f"{col}_T"] - merged[f"{col}_TR"]
        df_delta = merged[["year"] + sum_cols]

    taiga_sum = component_summary_from_yearly(df_taiga)
    trad_sum = component_summary_from_yearly(df_trad)
    taiga_total_pv = pv_total_from_yearly(df_taiga)
    trad_total_pv = pv_total_from_yearly(df_trad)
    delta_total_pv = taiga_total_pv - trad_total_pv
    effective_area = _get_effective_area(values)
    months = max(int(values["years"]) * 12, 1)
    area = max(float(effective_area), 1.0)
    leasing = _compute_leasing(values)

    payload = {
        "customer_name": "Demo Customer",
        "project_name": "Demo Project",
        "date_str": pd.Timestamp.now().strftime("%Y-%m-%d"),
        "params": {
            "years": int(values["years"]),
            "wacc": float(values["wacc"]),
            "shared_area_m2": float(values["area_m2"]),
            "area_m2": float(effective_area),
            "kwh_m2yr": float(values["kwh_m2yr"]),
            "elec_price": float(values["elec_price"]),
            "cycle_year": int(values["cycle_year"]),
        },
        "results": {
            "TCO_TRAD_PV": trad_total_pv,
            "TCO_TAIGA_PV": taiga_total_pv,
            "DIFF_TRAD_TAIGA": trad_total_pv - taiga_total_pv,
            "TAIGA_COST_M2_MONTH": taiga_total_pv / (area * months),
            "TRAD_COST_M2_MONTH": trad_total_pv / (area * months),
            "DELTA_COST_M2_MONTH": delta_total_pv / (area * months),
        },
        "taiga_forma": {
            "list_price": float(values["taiga_list_price"]),
            "units": int(values["taiga_total_qty"]),
            "effective_area_m2": float(effective_area),
            "occupancy_rate": float(values["taiga_occ_rate"]),
            "standby_share": float(values["taiga_standby"]),
            "commissioning_total": float(values["taiga_commissioning_cost_unit"]) * int(values["taiga_total_qty"]),
            "maintenance_total": float(values["taiga_maint_total_unit"]) * int(values["taiga_total_qty"]) * int(values["years"]),
            "end_of_life_cost": float(values["taiga_eol_cost"]),
        },
        "taiga_cycle": {
            "cycle_year": int(values["cycle_year"]),
        },
        "traditional_model": {
            "investment_per_m2": float(values["trad_price_per_m2"]),
            "room_qty": int(values["trad_room_qty"]),
            "run_fraction": float(values["trad_run_frac"]),
            "end_of_life_pct": float(values["trad_eol_pct"]),
            "commissioning_total": float(values["trad_commissioning_cost_unit"]) * int(values["trad_room_qty"]),
            "maintenance_total": float(values["trad_maint_total_unit"]) * int(values["trad_room_qty"]) * int(values["years"]),
        },
        "leasing": {
            "base_monthly": float(leasing["base_monthly"]),
            "monthly_with_buyback": float(leasing["monthly_with_buyback"]),
            "term_months": int(leasing["term_months"]),
            "buyback_year": int(values["lease_buyback_year"]),
        },
    }

    return {
        "values": values,
        "effective_area": effective_area,
        "taiga_sum": taiga_sum,
        "trad_sum": trad_sum,
        "taiga_total_pv": taiga_total_pv,
        "trad_total_pv": trad_total_pv,
        "delta_total_pv": delta_total_pv,
        "taiga_cost_m2_mo": taiga_total_pv / (area * months),
        "trad_cost_m2_mo": trad_total_pv / (area * months),
        "delta_cost_m2_mo": delta_total_pv / (area * months),
        "df_taiga": df_taiga,
        "df_trad": df_trad,
        "df_delta": df_delta,
        "pv_taiga": pivot_for_display(df_taiga).round(0),
        "pv_trad": pivot_for_display(df_trad).round(0),
        "pv_delta": pivot_for_display(df_delta).round(0) if df_delta is not None and not df_delta.empty else None,
        "leasing": leasing,
        "products_for_offer": _build_products_for_offer(values),
        "payload": payload,
    }


def _card(title: str, value: str, note: str = "", dark: bool = False):
    css_class = "taiga-card taiga-card--dark" if dark else "taiga-card"
    safe_note = note if str(note).strip() else "&nbsp;"
    st.markdown(
        f"""
        <div class="{css_class}">
          <div class="taiga-card__title">{title}</div>
          <div class="taiga-card__value">{value}</div>
          <div class="taiga-card__note">{safe_note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _go_to_step(step_idx: int):
    st.session_state.wizard_step = max(0, min(step_idx, len(WIZARD_STEPS) - 1))
    st.rerun()


def _render_header():
    current = int(st.session_state.wizard_step)
    title, subtitle = WIZARD_STEPS[current]
    st.markdown(
        f'<div class="taiga-hero">'
        f'<div class="taiga-brandbar">'
        f'<div class="taiga-brandbar__logo">'
        f'<img src="data:image/png;base64,{st.session_state.taiga_logo_b64}" alt="Taiga logo" />'
        f"</div>"
        f'<div class="taiga-brandbar__copy">'
        f'<div class="taiga-label">Taiga Spaces</div>'
        f'<p class="taiga-title">TAIGA SPACES: Lifecycle Calculator</p>'
        f'<p class="taiga-subtitle">Step {current + 1} of {len(WIZARD_STEPS)}. {title} - {subtitle}</p>'
        f"</div>"
        f"</div>"
        f"</div>",
        unsafe_allow_html=True,
    )
    cols = st.columns(len(WIZARD_STEPS))
    for idx, (label, _) in enumerate(WIZARD_STEPS):
        with cols[idx]:
            if st.button(f"{idx + 1}. {label}", key=f"wizard_step_{idx}", use_container_width=True, type="primary" if idx == current else "secondary"):
                _go_to_step(idx)


def _render_footer():
    current = int(st.session_state.wizard_step)
    left, mid, right = st.columns([1.2, 1.6, 1.2])
    with left:
        if st.button("< Previous", key="wizard_prev", use_container_width=True, disabled=current == 0):
            _go_to_step(current - 1)
    with mid:
        st.caption(f"Current step: {WIZARD_STEPS[current][0]}")
    with right:
        if st.button("Next >", key="wizard_next", use_container_width=True, disabled=current == len(WIZARD_STEPS) - 1):
            _go_to_step(current + 1)


def _render_project_basics():
    st.markdown("### Project Basics")
    st.caption("Set the shared assumptions used across the Traditional Model, Taiga Forma, lifecycle view and leasing view.")
    left, right = st.columns([0.52, 0.48], gap="large")
    with left:
        st.markdown('<div class="taiga-form-section">', unsafe_allow_html=True)
        st.number_input("Horizon (years)", 1, 50, key="years")
        _render_percent_input("WACC (%)", "wacc", "wacc_pct_ui")
        st.number_input("Area (m2)", 0.0, 1e9, key="area_m2", step=1.0)
        st.number_input("kWh / m2 / year", 0.0, 1e9, key="kwh_m2yr", step=1.0)
        st.number_input("Electricity price (EUR / kWh)", 0.0, 10.0, key="elec_price", step=0.01)
        st.markdown('</div>', unsafe_allow_html=True)
    with right:
        _card("Shared horizon", f"{int(st.session_state.years)} years", f"Base area {float(st.session_state.area_m2):,.2f} m2", dark=True)
        _card("Discount rate", f"{float(st.session_state.wacc) * 100:,.2f}%", "Applied across lifecycle calculations")
        _card("Energy intensity", f"{float(st.session_state.kwh_m2yr):,.2f} kWh / m2 / year", "")
        _card("Electricity price", _format_money(float(st.session_state.elec_price), 2), "Per kWh")


def _render_trad():
    st.markdown("### Traditional Model")
    st.caption("Define the conventional baseline and review the commercial and operational impact live on the right.")
    st.session_state.setdefault("trad_operational_defaults_sync", True)
    if st.session_state.get("trad_use_fitout_builder"):
        tuned_outputs = _compute_trad_fitout_outputs(float(max(st.session_state.area_m2, 10.0)), use_detailed=True)
        st.session_state.trad_price_per_m2 = float(tuned_outputs["per_m2"])
        st.session_state.trad_eol_pct = float(tuned_outputs["eol_pct"])
    operational_defaults = _compute_trad_operational_defaults(float(max(st.session_state.area_m2, 10.0)))
    if st.session_state.get("trad_operational_defaults_sync", True):
        st.session_state.trad_commissioning_cost_unit = float(operational_defaults["commissioning_cost_unit"])
        st.session_state.trad_maint_total_unit = float(operational_defaults["maint_total_unit"])
        st.session_state.trad_dt_rate = float(operational_defaults["dt_rate"])
        st.session_state.trad_dt_install_h_unit = float(operational_defaults["dt_install_h_unit"])
        st.session_state.trad_dt_maint_h_total_unit = float(operational_defaults["dt_maint_h_total_unit"])
    left, right = st.columns([0.52, 0.48], gap="large")
    with left:
        st.markdown('<div class="taiga-form-section">', unsafe_allow_html=True)
        st.number_input("Room count used for TRAD scaling", 0, 1_000_000, key="trad_room_qty", step=1)
        st.number_input("Investment price per m2 (TRAD) (EUR)", 0.0, 1e6, key="trad_price_per_m2", step=50.0)
        with st.expander("Open detailed CAT B / CAT C builder", expanded=bool(st.session_state.get("trad_use_fitout_builder", False))):
            _render_trad_fitout_builder(float(max(st.session_state.area_m2, 10.0)))
        st.slider("Run fraction (Traditional)", 0.0, 1.0, key="trad_run_frac")
        eol_pct_input = st.number_input("End-of-life (% of investment)", 0.0, 100.0, value=st.session_state.trad_eol_pct * 100, step=1.0)
        st.session_state.trad_eol_pct = float(eol_pct_input) / 100.0
        st.markdown(
            '<div class="taiga-subsection taiga-subsection--compact">'
            '<div class="taiga-subsection__title">Operational defaults</div>'
            '<div class="taiga-subsection__body">These default values update from the current benchmark and room geometry. You can switch the sync off and tune them manually.</div>',
            unsafe_allow_html=True,
        )
        st.checkbox("Keep operational defaults synced with benchmark", key="trad_operational_defaults_sync")
        st.markdown(
            f'<div class="taiga-inline-note">Current synced defaults: commissioning EUR {operational_defaults["commissioning_cost_unit"]:,.0f} per room, maintenance EUR {operational_defaults["maint_total_unit"]:,.0f} per room annually, downtime rate EUR {operational_defaults["dt_rate"]:,.1f} / h, install downtime {operational_defaults["dt_install_h_unit"]:,.1f} h per room and maintenance downtime {operational_defaults["dt_maint_h_total_unit"]:,.1f} h per room annually.</div>',
            unsafe_allow_html=True,
        )
        op_top = st.columns(2, gap="small")
        with op_top[0]:
            st.number_input("Commissioning cost per room (EUR)", 0.0, 1e9, key="trad_commissioning_cost_unit", step=10.0)
        with op_top[1]:
            st.number_input("Maintenance total per room annually (EUR)", 0.0, 1e12, key="trad_maint_total_unit", step=10.0)
        c1, c2, c3 = st.columns(3, gap="small")
        with c1:
            st.number_input("Downtime rate (EUR / h)", 0.0, 1e9, key="trad_dt_rate", step=0.1)
        with c2:
            st.number_input("Install downtime per room (h)", 0.0, 1e9, key="trad_dt_install_h_unit", step=0.1)
        with c3:
            st.number_input("Maintenance downtime per room annually (h)", 0.0, 1e9, key="trad_dt_maint_h_total_unit", step=0.1)
        st.number_input("Commissioning year", 0, 50, key="trad_commissioning_year")
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with right:
        effective_area = _get_effective_area(_snapshot_state())
        qty = int(st.session_state.trad_room_qty)
        _card("Traditional investment", _format_money(float(st.session_state.trad_price_per_m2) * effective_area), f"Area basis {effective_area:,.2f} m2", dark=True)
        _card("Rooms used", f"{qty:,}", "Scaling quantity")
        _card("Traditional maintenance total", _format_money(float(st.session_state.trad_maint_total_unit) * qty * int(st.session_state.years)), f"Across {int(st.session_state.years)} years")
        _card("Traditional commissioning total", _format_money(float(st.session_state.trad_commissioning_cost_unit) * qty), "")
        _card("Install downtime total", f"{float(st.session_state.trad_dt_install_h_unit) * qty:,.1f} h", "")
        _card("Maintenance downtime total", f"{float(st.session_state.trad_dt_maint_h_total_unit) * qty * int(st.session_state.years):,.1f} h", "")
        _card("End-of-life allowance", _format_money(float(st.session_state.trad_eol_pct) * float(st.session_state.trad_price_per_m2) * effective_area), "Percent of investment")


def _render_taiga(taiga_price_list_ui):
    st.markdown("### Taiga Forma")
    st.caption("Configure Taiga Forma first, then tune the operating and lifecycle assumptions. The summary updates continuously on the right.")
    left, right = st.columns([0.52, 0.48], gap="large")
    with left:
        st.markdown('<div class="taiga-form-section">', unsafe_allow_html=True)
        st.markdown(
            '<div class="taiga-subsection">'
            '<div class="taiga-subsection__title">Product selector</div>'
            '<div class="taiga-subsection__body">Choose Taiga Forma products, set quantities and let the total value update automatically.</div>',
            unsafe_allow_html=True,
        )
        _render_taiga_product_selector()
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown(
            '<div class="taiga-subsection">'
            '<div class="taiga-subsection__title">Commercial inputs</div>'
            '<div class="taiga-subsection__body">Use the current defaults as a baseline and adjust only when the project requires it.</div>',
            unsafe_allow_html=True,
        )
        st.markdown('<div class="taiga-inline-note">Taiga Forma list price is calculated automatically from the selected products.</div>', unsafe_allow_html=True)
        st.number_input(
            "Occupancy rate",
            min_value=0.0,
            max_value=1.0,
            key="taiga_occ_rate",
            step=0.01,
            format="%.2f",
        )
        st.number_input(
            "Standby share",
            min_value=0.0,
            max_value=1.0,
            key="taiga_standby",
            step=0.01,
            format="%.2f",
        )
        st.number_input(
            "Commissioning cost per unit (EUR)",
            min_value=0.0,
            max_value=1e12,
            key="taiga_commissioning_cost_unit",
            step=50.0,
        )
        st.number_input(
            "Maintenance per unit annually (EUR)",
            min_value=0.0,
            max_value=1e12,
            key="taiga_maint_total_unit",
            step=10.0,
        )
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown(
            '<div class="taiga-subsection">'
            '<div class="taiga-subsection__title">Downtime and lifecycle inputs</div>'
            '<div class="taiga-subsection__body">These values define commissioning timing, downtime cost and end-of-life treatment.</div>',
            unsafe_allow_html=True,
        )
        st.number_input("Downtime rate (EUR / h)", min_value=0.0, max_value=1e9, key="taiga_dt_rate", step=1.0)
        st.number_input("Install downtime per unit (h)", min_value=0.0, max_value=1e9, key="taiga_dt_install_h_unit", step=0.5)
        st.number_input("Maintenance downtime per unit annually (h)", min_value=0.0, max_value=1e9, key="taiga_dt_maint_h_total_unit", step=0.5)
        st.number_input("Commissioning year", min_value=0, max_value=50, key="taiga_commissioning_year", step=1)
        st.number_input("End-of-life cost (EUR)", min_value=0.0, max_value=1e12, key="taiga_eol_cost", step=100.0)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown(
            '<div class="taiga-subsection">'
            '<div class="taiga-subsection__title">Advanced price list tools</div>'
            '<div class="taiga-subsection__body">Use these only when you need to replace or edit the underlying Taiga Forma price list.</div>',
            unsafe_allow_html=True,
        )
        with st.expander("Open advanced tools", expanded=False):
            st.checkbox("Use selected product area as effective project area", key="override_area", disabled=int(st.session_state.get("taiga_total_qty", 0)) == 0)
            _load_price_list_upload(key="taiga_price_upload_advanced", label="Upload replacement price list (xlsx / csv)")
            st.session_state.price_df = st.data_editor(
                st.session_state.price_df,
                width="stretch",
                num_rows="dynamic",
                key="price_table_editor_advanced",
                column_config={"qty": st.column_config.NumberColumn("qty", min_value=0, step=1)},
            ).copy()
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    qty = int(st.session_state.get("taiga_total_qty", 0))
    with right:
        effective_area = _get_effective_area(_snapshot_state())
        taiga_maint_total = float(st.session_state.taiga_maint_total_unit) * qty * int(st.session_state.years)
        taiga_comm_total = float(st.session_state.taiga_commissioning_cost_unit) * qty
        taiga_install_dt = float(st.session_state.taiga_dt_install_h_unit) * qty
        taiga_maint_dt = float(st.session_state.taiga_dt_maint_h_total_unit) * qty * int(st.session_state.years)

        _card("Taiga Forma total value", _format_money(float(st.session_state.taiga_list_price)), "Calculated from the selected products", dark=True)
        _card("Selected Taiga Forma units", f"{qty:,}", "From the current product mix")
        _card("Effective project area", f"{effective_area:,.2f} m2", "Product area override respected")
        _render_taiga_configuration_card()
        _card("Taiga Forma commissioning total", _format_money(taiga_comm_total), "Scaled by selected units")
        _card("Taiga Forma maintenance total", _format_money(taiga_maint_total), f"Across {int(st.session_state.years)} years")
        _card("End-of-life cost", _format_money(float(st.session_state.taiga_eol_cost)), "Nominal total input")
        _card("Install downtime total", f"{taiga_install_dt:,.1f} h", "")
        _card("Maintenance downtime total", f"{taiga_maint_dt:,.1f} h", "")


def _render_buyback():
    st.markdown("### Taiga Cycle")
    st.caption("Set the Taiga Cycle year and maintain the value curve used in the lifecycle model.")
    left, right = st.columns([0.52, 0.48], gap="large")
    with left:
        st.markdown('<div class="taiga-form-section">', unsafe_allow_html=True)
        st.number_input("Cycle (buyback) year", 0, 50, key="cycle_year")
        st.caption("Edit or add rows. Values can be entered as 70 or 0.70.")
        with st.expander("Open buyback value table", expanded=False):
            st.session_state.cycle_df = st.data_editor(st.session_state.cycle_df, num_rows="dynamic", width="stretch", key="cycle_editor")
        st.markdown('</div>', unsafe_allow_html=True)
    with right:
        row = st.session_state.cycle_df.loc[st.session_state.cycle_df["year"] == int(st.session_state.cycle_year)]
        pct = 0.0
        if not row.empty:
            raw = float(row["value_pct"].iloc[0])
            pct = raw / 100.0 if abs(raw) > 1.0 else raw
        buyback_price = float(st.session_state.taiga_list_price) * pct
        _card("Selected Taiga Cycle year", f"{int(st.session_state.cycle_year)}", "Set to 0 to disable", dark=True)
        _card("Taiga Cycle value", f"{pct * 100:,.1f}%", "Read from the current cycle table")
        _card("Taiga Forma list price", _format_money(float(st.session_state.taiga_list_price)), "Used as the cycle base")
        _card("Taiga Cycle buyback price", _format_money(buyback_price), "Nominal value before discounting")


def _render_leasing():
    st.markdown("### Leasing")
    st.caption("Keep financing assumptions together so the monthly payment view stays easy to review.")
    left, right = st.columns([0.52, 0.48], gap="large")
    with left:
        st.markdown('<div class="taiga-form-section">', unsafe_allow_html=True)
        st.number_input("Leasing term (years)", 1, 15, key="lease_term_years", step=1)
        _render_percent_input("WACC annual (%)", "lease_wacc_annual", "lease_wacc_pct_ui")
        st.number_input("Base price (EUR)", 0.0, 1e12, key="lease_base_price", step=100.0)
        st.number_input("Buyback year (0 = none)", 0, 50, key="lease_buyback_year", step=1)
        with st.expander("Open monthly factor table", expanded=False):
            st.session_state.lease_factors_df = st.data_editor(
                st.session_state.lease_factors_df,
                width="stretch",
                num_rows="dynamic",
                key="lease_factors_editor",
                column_config={
                    "term_years": st.column_config.NumberColumn("term_years", min_value=1, step=1),
                    "monthly_factor": st.column_config.NumberColumn("monthly_factor", help="Enter as % per month (e.g. 1.55) or decimal (0.0155)."),
                },
            )
        st.markdown('</div>', unsafe_allow_html=True)
    leasing = _compute_leasing(_snapshot_state())
    st.session_state["lease_monthly_price"] = float(leasing["monthly_with_buyback"])
    st.session_state["lease_term_months"] = int(leasing["term_months"])
    with right:
        _card("Monthly factor", f"{leasing['monthly_factor'] * 100:.2f}%", "", dark=True)
        _card("Base monthly payment", _format_money(leasing["base_monthly"]), "")
        _card("Monthly payment with buyback", _format_money(leasing["monthly_with_buyback"]), f"Delta {leasing['monthly_with_buyback'] - leasing['base_monthly']:,.0f}")
        _card("Leasing term", f"{int(st.session_state.lease_term_years)} years", f"{leasing['term_months']} months")


def _render_total_chart(results: dict):
    scale_divisor = 1000.0
    chart_df = pd.DataFrame([
        {"Model": "Taiga", "Total PV": results["taiga_total_pv"], "Display Value": results["taiga_total_pv"] / scale_divisor},
        {"Model": "TRAD", "Total PV": results["trad_total_pv"], "Display Value": results["trad_total_pv"] / scale_divisor},
        {"Model": "Delta", "Total PV": results["delta_total_pv"], "Display Value": results["delta_total_pv"] / scale_divisor},
    ])
    chart = (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Model:N", sort=["Taiga", "TRAD", "Delta"]),
            y=alt.Y("Display Value:Q", title="Present value (kEUR)"),
            color=alt.Color("Model:N", scale=alt.Scale(domain=["Taiga", "TRAD", "Delta"], range=["#1E4D35", "#C47C2A", "#5C5B56"])),
            tooltip=[
                "Model",
                alt.Tooltip("Total PV:Q", title="Present value (EUR)", format=",.0f"),
                alt.Tooltip("Display Value:Q", title="Present value (kEUR)", format=",.1f"),
            ],
            text=alt.Text("Display Value:Q", format=",.1f"),
        )
        .properties(height=320)
    )
    text = (
        alt.Chart(chart_df)
        .mark_text(dy=-10, fontSize=12, fontWeight="bold")
        .encode(
            x=alt.X("Model:N", sort=["Taiga", "TRAD", "Delta"]),
            y=alt.Y("Display Value:Q"),
            text=alt.Text("Display Value:Q", format=",.1f"),
            color=alt.value("#5C5B56"),
        )
        .properties(height=320)
    )
    st.altair_chart(chart + text, use_container_width=True)


def _render_component_chart(results: dict):
    label_map = {
        "acquisition_pv": "Acquisition",
        "buyback_pv": "Buyback",
        "commissioning_pv": "Commissioning",
        "energy_cost_pv": "Energy",
        "maintenance_pv": "Maintenance",
        "downtime_pv": "Downtime",
        "eol_pv": "End-of-life",
    }
    rows = []
    for key, label in label_map.items():
        rows.append({"Component": label, "Model": "Taiga", "Value": float(results["taiga_sum"].get(key, 0.0)), "Display Value": float(results["taiga_sum"].get(key, 0.0)) / 1000.0})
        rows.append({"Component": label, "Model": "TRAD", "Value": float(results["trad_sum"].get(key, 0.0)), "Display Value": float(results["trad_sum"].get(key, 0.0)) / 1000.0})
    chart_df = pd.DataFrame(rows)
    chart = (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
        .encode(
            x=alt.X("Component:N", sort=list(label_map.values()), title="Component"),
            xOffset=alt.XOffset("Model:N"),
            y=alt.Y("Display Value:Q", title="Present value (kEUR)"),
            color=alt.Color("Model:N", scale=alt.Scale(domain=["Taiga", "TRAD"], range=["#1E4D35", "#C47C2A"])),
            tooltip=[
                "Component",
                "Model",
                alt.Tooltip("Value:Q", title="Present value (EUR)", format=",.0f"),
                alt.Tooltip("Display Value:Q", title="Present value (kEUR)", format=",.1f"),
            ],
        )
        .properties(height=320)
    )
    text = (
        alt.Chart(chart_df)
        .mark_text(dy=-8, fontSize=11)
        .encode(
            x=alt.X("Component:N", sort=list(label_map.values())),
            xOffset=alt.XOffset("Model:N"),
            y=alt.Y("Display Value:Q"),
            detail="Model:N",
            text=alt.Text("Display Value:Q", format=",.1f"),
            color=alt.value("#5C5B56"),
        )
        .properties(height=320)
    )
    st.altair_chart(chart + text, use_container_width=True)


def _render_scenario_chart(ctx):
    st.markdown("#### TRAD price scenario view")
    st.caption("Plot lifecycle cost over time while varying the Traditional investment price per m2.")
    base_values = _snapshot_state()
    mode_cols = st.columns([0.24, 0.76], gap="small")
    with mode_cols[0]:
        mode = st.selectbox("Scenario mode", ["Sensitivity", "Custom values"], key="trad_price_scenario_mode")

    scenario_prices = []
    base_price = float(base_values["trad_price_per_m2"])
    if mode == "Sensitivity":
        s1, s2, s3 = st.columns(3, gap="small")
        with s1:
            band_pct = st.number_input("Band (+/- %)", min_value=5.0, max_value=100.0, value=20.0, step=5.0, key="trad_price_sensitivity_band")
        with s2:
            points = st.number_input("Scenarios", min_value=3, max_value=9, value=5, step=1, key="trad_price_sensitivity_points")
        with s3:
            st.number_input("Baseline TRAD EUR / m2", min_value=0.0, max_value=1e6, value=base_price, step=50.0, key="trad_price_sensitivity_baseline", disabled=True)
        low = base_price * (1.0 - (band_pct / 100.0))
        high = base_price * (1.0 + (band_pct / 100.0))
        step_size = (high - low) / max(int(points) - 1, 1)
        scenario_prices = [round(low + (idx * step_size), 0) for idx in range(int(points))]
    else:
        raw = st.text_input("Custom TRAD EUR / m2 values", value=f"{max(base_price - 500, 0):.0f}, {base_price:.0f}, {base_price + 500:.0f}", key="trad_price_custom_values")
        scenario_prices = []
        for item in raw.split(","):
            item = item.strip()
            if not item:
                continue
            try:
                scenario_prices.append(float(item))
            except ValueError:
                continue
        scenario_prices = sorted(set([value for value in scenario_prices if value >= 0]))

    if not scenario_prices:
        st.info("Add at least one valid TRAD EUR / m2 scenario value.")
        return

    rows = []
    taiga_added = False
    for scenario_price in scenario_prices:
        values = dict(base_values)
        values["trad_price_per_m2"] = float(scenario_price)
        result = _compute_results(values, ctx)

        trad_df = result["df_trad"].copy()
        trad_value_cols = [col for col in trad_df.columns if col != "year" and pd.api.types.is_numeric_dtype(trad_df[col])]
        trad_df["year_cost"] = trad_df[trad_value_cols].sum(axis=1)
        trad_df["cumulative_cost"] = trad_df["year_cost"].cumsum()
        for row in trad_df[["year", "cumulative_cost"]].itertuples(index=False):
            rows.append({
                "Year": int(row.year),
                "Cost": float(row.cumulative_cost),
                "Series": f"Traditional EUR {scenario_price:,.0f} / m2",
                "Model": "Traditional",
            })

        if not taiga_added:
            taiga_df = result["df_taiga"].copy()
            taiga_value_cols = [col for col in taiga_df.columns if col != "year" and pd.api.types.is_numeric_dtype(taiga_df[col])]
            taiga_df["year_cost"] = taiga_df[taiga_value_cols].sum(axis=1)
            taiga_df["cumulative_cost"] = taiga_df["year_cost"].cumsum()
            for row in taiga_df[["year", "cumulative_cost"]].itertuples(index=False):
                rows.append({
                    "Year": int(row.year),
                    "Cost": float(row.cumulative_cost),
                    "Series": "Taiga baseline",
                    "Model": "Taiga",
                })
            taiga_added = True

    chart_df = pd.DataFrame(rows)
    chart_df["Cost kEUR"] = chart_df["Cost"] / 1000.0
    color_scale = alt.Scale(
        domain=["Taiga baseline"] + [f"Traditional EUR {value:,.0f} / m2" for value in scenario_prices],
        range=["#1E4D35", "#C47C2A", "#B57D32", "#C99556", "#D8AF7A", "#E5C89E", "#EEDABD", "#F3E6D1", "#F7EFE3", "#FBF6ED"][: len(scenario_prices) + 1],
    )
    chart = (
        alt.Chart(chart_df)
        .mark_line(point=True)
        .encode(
            x=alt.X("Year:Q", title="Time (year)"),
            y=alt.Y("Cost kEUR:Q", title="Cumulative lifecycle cost (kEUR)"),
            color=alt.Color("Series:N", scale=color_scale),
            tooltip=[
                alt.Tooltip("Series:N", title="Scenario"),
                alt.Tooltip("Year:Q", title="Year"),
                alt.Tooltip("Cost:Q", title="Cost (EUR)", format=",.0f"),
                alt.Tooltip("Cost kEUR:Q", title="Cost (kEUR)", format=",.1f"),
            ],
        )
        .properties(height=360)
    )
    st.altair_chart(chart, use_container_width=True)


def _render_summary(ctx):
    results = _compute_results(_snapshot_state(), ctx)
    st.session_state["lease_monthly_price"] = float(results["leasing"]["monthly_with_buyback"])
    st.session_state["lease_term_months"] = int(results["leasing"]["term_months"])
    st.markdown("### Executive Summary")
    st.caption("Compare the current baseline here, then return to earlier steps whenever you want to refine assumptions.")
    top = st.columns(4)
    with top[0]:
        _card("Taiga total present value", _format_money(results["taiga_total_pv"]), f"{results['effective_area']:,.2f} m2 basis", dark=True)
    with top[1]:
        _card("Traditional total present value", _format_money(results["trad_total_pv"]), "Traditional baseline")
    with top[2]:
        _card("Delta total present value", _format_money(results["delta_total_pv"]), "Taiga minus Traditional")
    with top[3]:
        _card("Monthly leasing payment", _format_money(results["leasing"]["monthly_with_buyback"]), f"{results['leasing']['term_months']} months")
    lower = st.columns(4)
    with lower[0]:
        _card("Taiga average lifecycle cost (EUR / m2 / month)", f"EUR {results['taiga_cost_m2_mo']:,.2f}", "Discounted lifecycle basis")
    with lower[1]:
        _card("Traditional average lifecycle cost (EUR / m2 / month)", f"EUR {results['trad_cost_m2_mo']:,.2f}", "Discounted lifecycle basis")
    with lower[2]:
        _card("Delta average lifecycle cost (EUR / m2 / month)", f"EUR {results['delta_cost_m2_mo']:,.2f}", "Taiga minus Traditional")
    with lower[3]:
        _card("Effective project area", f"{results['effective_area']:,.2f} m2", "Area override respected")
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown("#### Total value comparison")
        _render_total_chart(results)
    with c2:
        st.markdown("#### Component comparison")
        _render_component_chart(results)
    edit_cols = st.columns(5)
    for idx, (label, _) in enumerate(WIZARD_STEPS[:5]):
        with edit_cols[idx]:
            if st.button(f"Edit {label}", key=f"edit_step_{idx}", use_container_width=True):
                _go_to_step(idx)
    _render_scenario_chart(ctx)


def _render_reporting(ctx):
    results = _compute_results(_snapshot_state(), ctx)
    st.session_state["lease_monthly_price"] = float(results["leasing"]["monthly_with_buyback"])
    st.session_state["lease_term_months"] = int(results["leasing"]["term_months"])
    st.markdown("### Lifecycle Costs and Reporting")
    st.caption("Review the detailed lifecycle outputs and export the current Taiga Forma and Traditional Model summary.")
    keys = sorted(set(results["taiga_sum"].keys()) | set(results["trad_sum"].keys()))
    delta = {k: results["taiga_sum"].get(k, 0.0) - results["trad_sum"].get(k, 0.0) for k in keys}

    top_tables = st.columns(2, gap="large")
    with top_tables[0]:
        _render_report_table(pd.DataFrame([results["taiga_sum"]]).T.rename(columns={0: "Taiga"}), "Taiga summary", index=True)
    with top_tables[1]:
        _render_report_table(pd.DataFrame([results["trad_sum"]]).T.rename(columns={0: "Traditional"}), "Traditional summary", index=True)

    mid_tables = st.columns(2, gap="large")
    with mid_tables[0]:
        _render_report_table(results["pv_taiga"], "Taiga lifecycle breakdown", index=True)
    with mid_tables[1]:
        _render_report_table(results["pv_trad"], "Traditional lifecycle breakdown", index=True)

    bottom_tables = st.columns(2, gap="large")
    with bottom_tables[0]:
        _render_report_table(pd.DataFrame([delta]).T.rename(columns={0: "Delta"}), "Delta summary", index=True)
    with bottom_tables[1]:
        _render_report_table(results["leasing"]["pivot"], "Leasing breakdown", index=True)

    ctx["reload_module"](ctx["proposal_doc"])
    ctx["reload_module"](ctx["offer_doc"])
    doc_bytes = ctx["proposal_doc"].generate_proposal_doc(
        payload=results["payload"],
        df_pivot_taiga=results["pv_taiga"],
        df_pivot_trad=results["pv_trad"],
        df_pivot_delta=results["pv_delta"],
        locale="fi_FI",
        logo_path=str(ctx["app_dir"] / "logo.PNG"),
    )
    offer_bytes = ctx["offer_doc"].generate_offer_doc(
        payload=results["payload"],
        products_df=results["products_for_offer"],
        leasing_info={
            "monthly_price_base": results["leasing"]["base_monthly"],
            "monthly_price_with_buyback": results["leasing"]["monthly_with_buyback"],
            "term_months": results["leasing"]["term_months"],
            "buyback_year": results["values"]["lease_buyback_year"],
        },
        trad_summary={
            "taiga_pv": results["taiga_total_pv"],
            "trad_pv": results["trad_total_pv"],
            "delta_pv": results["delta_total_pv"],
            "taiga_cost_m2_mo": results["taiga_cost_m2_mo"],
            "trad_cost_m2_mo": results["trad_cost_m2_mo"],
            "effective_area": results["effective_area"],
            "taiga_list_price": results["values"]["taiga_list_price"],
            "cycle_year": results["values"]["cycle_year"],
        },
        logo_path=str(ctx["app_dir"] / "logo.PNG"),
    )
    b1, b2 = st.columns(2)
    with b1:
        st.download_button("Download Word summary", doc_bytes, "TCO_Summary.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with b2:
        st.download_button("Download offer document", offer_bytes, "Taiga_Offer.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)


def render_app(**ctx):
    _sync_ui_state()
    app_dir = Path(ctx["app_dir"])
    logo_b64 = base64.b64encode((app_dir / "logo.PNG").read_bytes()).decode()
    st.session_state.taiga_logo_b64 = logo_b64
    css = ctx["base_css"].replace("URL_LOGO_PLACEHOLDER", f"data:image/png;base64,{logo_b64}") + TAIGA_EXTRA_CSS
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)

    _render_header()
    step = int(st.session_state.wizard_step)
    if step == 0:
        _render_project_basics()
    elif step == 1:
        _render_trad()
    elif step == 2:
        _render_taiga(ctx["taiga_price_list_ui"])
    elif step == 3:
        _render_buyback()
    elif step == 4:
        _render_leasing()
    elif step == 5:
        _render_summary(ctx)
    elif step == 6:
        _render_reporting(ctx)
    _render_footer()
