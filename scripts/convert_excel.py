#!/usr/bin/env python3
"""CS WIND Equipment-Based Productivity dashboard generator.

Reads Raw Data_Global Equipment-Based Productivity.xlsx and generates:
- index.html
- dashboard_template.html

The layout is the old style:
View / Weekly / Monthly / Year / Week only.
No "Factory for Trend" selector in the top area.
"""

import json
import math
import sys
from pathlib import Path
from datetime import date

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = BASE_DIR / "Raw Data_Global Equipment-Based Productivity.xlsx"
TEMPLATE_PATH = BASE_DIR / "dashboard_template.html"
OUTPUT_PATH = BASE_DIR / "index.html"

FACTORIES_ORDER = ["VN #1", "VN #2", "TW", "CN", "TR #1", "TR #2", "AM", "PT On", "PT Off"]

EQUIPMENT_MAP = {
    "bending": "Roll Bending Machine",
    "lw": "L/W Machine",
    "cw": "C/W Machine",
    "growing": "Growing Line",
    "paintBooth": "Paint Booth",
    "paintLine": "Paint Line",
}

PRODUCTION_MAP = {
    "bending": "Bending",
    "lw": "L/W",
    "cw": "C/W",
    "btgt": "BT GT",
    "wtgt": "WT GT",
}

MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def safe_number(value):
    try:
        if pd.isna(value):
            return 0.0
        if isinstance(value, str):
            value = value.replace(",", "").strip()
            if value in ("", "-", "#N/A"):
                return 0.0
        num = float(value)
        if math.isnan(num) or math.isinf(num):
            return 0.0
        return num
    except Exception:
        return 0.0


def week_to_month(year, week):
    try:
        y = int(float(year))
        w = max(1, min(int(float(week)), 53))
        d = date.fromisocalendar(y, w, 1)
        return MONTH_NAMES[d.month - 1], f"{y}-{MONTH_NAMES[d.month - 1]}", y * 100 + d.month
    except Exception:
        return "", "", 0


def normalize_dataframe(excel_path):
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    try:
        df = pd.read_excel(excel_path, sheet_name="GitHub_Export", engine="openpyxl")
        print(f"Using GitHub_Export sheet: {len(df)} rows")
    except Exception as exc:
        print(f"GitHub_Export unavailable, fallback to Raw(1): {exc}")
        df = pd.read_excel(excel_path, sheet_name="Raw(1)", engine="openpyxl")

    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    required = ["Year", "Week", "Factory", "Type", "Category", "Q'ty"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"GitHub_Export missing columns {missing}; fallback to Raw(1)")
        df = pd.read_excel(excel_path, sheet_name="Raw(1)", engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Required columns missing: {missing}")

    keep_cols = required + [c for c in ["Month", "YearMonth", "Month_Sort"] if c in df.columns]
    df = df[keep_cols].copy()

    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["Week"] = pd.to_numeric(df["Week"], errors="coerce")
    df["Q'ty"] = pd.to_numeric(df["Q'ty"], errors="coerce").fillna(0)

    df = df.dropna(subset=["Year", "Week", "Factory", "Type", "Category"])
    df["Year"] = df["Year"].astype(int)
    df["Week"] = df["Week"].astype(int)
    df["Factory"] = df["Factory"].astype(str).str.strip()
    df["Type"] = df["Type"].astype(str).str.strip()
    df["Category"] = df["Category"].astype(str).str.strip()

    calculated = df.apply(lambda r: week_to_month(r["Year"], r["Week"]), axis=1)

    if "Month" not in df.columns:
        df["Month"] = calculated.apply(lambda x: x[0])
    else:
        df["Month"] = df["Month"].fillna("").astype(str).str.strip()

    if "YearMonth" not in df.columns:
        df["YearMonth"] = calculated.apply(lambda x: x[1])
    else:
        df["YearMonth"] = df["YearMonth"].fillna("").astype(str).str.strip()

    if "Month_Sort" not in df.columns:
        df["Month_Sort"] = calculated.apply(lambda x: x[2]).astype(int)
    else:
        df["Month_Sort"] = pd.to_numeric(df["Month_Sort"], errors="coerce").fillna(0).astype(int)

    mask = (df["Month"] == "") | (df["YearMonth"] == "") | (df["Month_Sort"] == 0)
    if mask.any():
        df.loc[mask, "Month"] = calculated[mask].apply(lambda x: x[0])
        df.loc[mask, "YearMonth"] = calculated[mask].apply(lambda x: x[1])
        df.loc[mask, "Month_Sort"] = calculated[mask].apply(lambda x: x[2]).astype(int)

    df = df[(df["Year"] > 0) & (df["Week"] > 0) & (df["Factory"] != "")]
    print(f"Rows: {len(df):,}")
    print(f"Years: {sorted(df['Year'].unique().tolist())}")
    print(f"Weeks: WK{df['Week'].min():02d} ~ WK{df['Week'].max():02d}")
    return df


def ordered_factories(df):
    found = list(dict.fromkeys(df["Factory"].dropna().astype(str).str.strip().tolist()))
    ordered = [f for f in FACTORIES_ORDER if f in found]
    ordered += [f for f in sorted(found) if f not in ordered]
    return ordered


def get_val(df_fac, type_name, category):
    value = df_fac[(df_fac["Type"] == type_name) & (df_fac["Category"] == category)]["Q'ty"].sum()
    return safe_number(value)


def build_period_payload(df_period):
    factories = ordered_factories(df_period)
    equipment = {}
    production = {}

    for factory in factories:
        df_fac = df_period[df_period["Factory"] == factory]

        equipment[factory] = {
            key: get_val(df_fac, "Machine", category)
            for key, category in EQUIPMENT_MAP.items()
        }

        production[factory] = {
            key: get_val(df_fac, "Performance", category)
            for key, category in PRODUCTION_MAP.items()
        }

    return {
        "equipment": equipment,
        "production": production,
    }


def convert_weekly(df):
    raw_data = {}

    for year in sorted(df["Year"].unique()):
        year_key = f"{int(year)}Y"
        raw_data[year_key] = {}
        df_year = df[df["Year"] == year]

        for week in sorted(df_year["Week"].unique()):
            week_key = f"WK{int(week):02d}"
            raw_data[year_key][week_key] = build_period_payload(df_year[df_year["Week"] == week])

    return raw_data


def convert_monthly(df):
    monthly_data = {}

    for year in sorted(df["Year"].unique()):
        year_key = f"{int(year)}Y"
        monthly_data[year_key] = {}
        df_year = df[df["Year"] == year].sort_values("Month_Sort")

        months = (
            df_year[["YearMonth", "Month_Sort"]]
            .drop_duplicates()
            .sort_values("Month_Sort")
        )

        for _, row in months.iterrows():
            period_key = str(row["YearMonth"])
            if period_key and period_key != "nan":
                monthly_data[year_key][period_key] = build_period_payload(
                    df_year[df_year["YearMonth"] == period_key]
                )

    return monthly_data


def build_index_html(raw_data, monthly_data):
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_PATH}")

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    if "%%RAWDATA_PLACEHOLDER%%" not in template:
        raise ValueError("dashboard_template.html does not contain %%RAWDATA_PLACEHOLDER%%")

    js_data = (
        "const rawData = "
        + json.dumps(raw_data, ensure_ascii=False, indent=2)
        + ";\n"
        + "const monthlyData = "
        + json.dumps(monthly_data, ensure_ascii=False, indent=2)
        + ";"
    )

    output_html = template.replace("%%RAWDATA_PLACEHOLDER%%", js_data)
    OUTPUT_PATH.write_text(output_html, encoding="utf-8")

    # Keep template unchanged; only index.html is generated.
    print(f"Generated index.html: {len(output_html):,} bytes")


def main():
    df = normalize_dataframe(EXCEL_PATH)
    raw_data = convert_weekly(df)
    monthly_data = convert_monthly(df)
    build_index_html(raw_data, monthly_data)

    latest_year = sorted(raw_data.keys())[-1]
    latest_week = sorted(raw_data[latest_year].keys())[-1]
    print(f"Latest weekly period: {latest_year} {latest_week}")
    print("Done.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
