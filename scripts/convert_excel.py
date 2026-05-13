#!/usr/bin/env python3
"""CS WIND old-style dashboard generator.
Reads Raw Data_Global Equipment-Based Productivity.xlsx and generates index.html
using dashboard_template.html. Keeps the previous UI: View / Weekly / Monthly / Year / Week only,
without Factory for Trend cards.
"""
import json
import sys
from pathlib import Path
import pandas as pd

REQUIRED = ['Year', 'Week', 'Factory', 'Type', 'Category', "Q'ty"]

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
    for col in REQUIRED:
        if col not in df.columns:
            raise ValueError(f"Required column missing: {col}")
    keep = REQUIRED + [c for c in ['Month', 'YearMonth', 'Month_Sort'] if c in df.columns]
    df = df[keep]
    df['Year'] = pd.to_numeric(df['Year'], errors='coerce')
    df['Week'] = pd.to_numeric(df['Week'], errors='coerce')
    df["Q'ty"] = pd.to_numeric(df["Q'ty"], errors='coerce').fillna(0)
    df = df.dropna(subset=['Year','Week','Factory','Type','Category'])
    df['Year'] = df['Year'].astype(int)
    df['Week'] = df['Week'].astype(int)
    return df

def add_month_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if 'YearMonth' in df.columns and 'Month_Sort' in df.columns:
        return df
    base_date = pd.to_datetime(df['Year'].astype(str) + '-01-01', errors='coerce') + pd.to_timedelta((df['Week'] - 1) * 7, unit='D')
    df['Month'] = base_date.dt.strftime('%b')
    df['YearMonth'] = base_date.dt.strftime('%Y-%b')
    df['Month_Sort'] = base_date.dt.year * 100 + base_date.dt.month
    return df

def read_source(excel_path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(excel_path, sheet_name='GitHub_Export', engine='openpyxl')
        df = clean_df(df)
        if len(df) > 0:
            print(f"Using GitHub_Export: {len(df)} rows")
            return add_month_columns(df)
    except Exception as e:
        print(f"GitHub_Export unavailable, fallback to Raw(1): {e}")
    df = pd.read_excel(excel_path, sheet_name='Raw(1)', engine='openpyxl')
    df = clean_df(df)
    print(f"Using Raw(1): {len(df)} rows")
    return add_month_columns(df)

def get_val(df_fac: pd.DataFrame, type_: str, cat: str) -> float:
    v = df_fac[(df_fac['Type'] == type_) & (df_fac['Category'] == cat)]["Q'ty"].sum()
    return float(v) if not pd.isna(v) else 0.0

def convert_by_period(df: pd.DataFrame, period_col: str) -> dict:
    raw_data = {}
    for year in sorted(df['Year'].dropna().unique()):
        year_key = f"{int(year)}Y"
        raw_data[year_key] = {}
        df_year = df[df['Year'] == year]
        if period_col == 'Week':
            periods = sorted(df_year['Week'].dropna().unique())
        else:
            temp = df_year[['YearMonth','Month_Sort']].drop_duplicates().sort_values('Month_Sort')
            periods = temp['YearMonth'].tolist()
        for period in periods:
            if period_col == 'Week':
                period_key = f"WK{int(period):02d}"
                df_period = df_year[df_year['Week'] == period]
            else:
                period_key = str(period)
                df_period = df_year[df_year['YearMonth'] == period]
            equipment = {}
            production = {}
            for factory in df_period['Factory'].dropna().unique():
                df_fac = df_period[df_period['Factory'] == factory]
                equipment[factory] = {
                    'bending': get_val(df_fac, 'Machine', 'Roll Bending Machine'),
                    'lw': get_val(df_fac, 'Machine', 'L/W Machine'),
                    'cw': get_val(df_fac, 'Machine', 'C/W Machine'),
                    'growing': get_val(df_fac, 'Machine', 'Growing Line'),
                    'paintBooth': get_val(df_fac, 'Machine', 'Paint Booth'),
                    'paintLine': get_val(df_fac, 'Machine', 'Paint Line'),
                }
                production[factory] = {
                    'bending': get_val(df_fac, 'Performance', 'Bending'),
                    'lw': get_val(df_fac, 'Performance', 'L/W'),
                    'cw': get_val(df_fac, 'Performance', 'C/W'),
                    'btgt': get_val(df_fac, 'Performance', 'BT GT'),
                    'wtgt': get_val(df_fac, 'Performance', 'WT GT'),
                }
            raw_data[year_key][period_key] = {'equipment': equipment, 'production': production}
    return raw_data

def build_index_html(weekly_raw_data: dict, monthly_raw_data: dict, template_path: Path, output_path: Path) -> None:
    template = template_path.read_text(encoding='utf-8')
    js_data = 'const rawData = ' + json.dumps(weekly_raw_data, ensure_ascii=False, indent=2) + ';\n'
    js_data += 'const monthlyData = ' + json.dumps(monthly_raw_data, ensure_ascii=False, indent=2) + ';'
    if '%%RAWDATA_PLACEHOLDER%%' not in template:
        print('dashboard_template.html placeholder missing: %%RAWDATA_PLACEHOLDER%%')
        sys.exit(1)
    output_path.write_text(template.replace('%%RAWDATA_PLACEHOLDER%%', js_data), encoding='utf-8')
    print(f'Generated {output_path}')

def main() -> None:
    base_dir = Path(__file__).parent.parent
    excel_path = base_dir / 'Raw Data_Global Equipment-Based Productivity.xlsx'
    template_path = base_dir / 'dashboard_template.html'
    output_path = base_dir / 'index.html'
    if not excel_path.exists():
        print(f'Excel file not found: {excel_path}')
        sys.exit(1)
    if not template_path.exists():
        print(f'Template file not found: {template_path}')
        sys.exit(1)
    df = read_source(excel_path)
    weekly_raw_data = convert_by_period(df, 'Week')
    monthly_raw_data = convert_by_period(df, 'YearMonth')
    build_index_html(weekly_raw_data, monthly_raw_data, template_path, output_path)
    print('Done. Old-style Weekly/Monthly dashboard generated without Factory for Trend.')

if __name__ == '__main__':
    main()
