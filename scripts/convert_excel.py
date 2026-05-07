#!/usr/bin/env python3
"""
CS WIND Dashboard - Excel to HTML Auto Converter
- Reads Raw Data_Global Equipment-Based Productivity.xlsx / Raw(1)
- Builds both Weekly and Monthly dashboard data
- Generates index.html from dashboard_template.html
"""
import json
import sys
from pathlib import Path
from datetime import date
import pandas as pd

FACTORIES_ORDER = ['VN #1', 'VN #2', 'TW', 'CN', 'TR #1', 'TR #2', 'AM', 'PT On', 'PT Off']

EQUIPMENT_MAP = {
    'bending': 'Roll Bending Machine',
    'lw': 'L/W Machine',
    'cw': 'C/W Machine',
    'growing': 'Growing Line',
    'paintBooth': 'Paint Booth',
    'paintLine': 'Paint Line',
}

PRODUCTION_MAP = {
    'bending': 'Bending',
    'lw': 'L/W',
    'cw': 'C/W',
    'btgt': 'BT GT',
    'wtgt': 'WT GT',
}

MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']


def safe_number(value):
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def get_val(df_fac, type_, category):
    v = df_fac[(df_fac['Type'] == type_) & (df_fac['Category'] == category)]["Q'ty"].sum()
    return safe_number(v)


def week_to_month(year, week):
    try:
        y = int(float(year))
        w = int(float(week))
        d = date.fromisocalendar(y, max(1, min(w, 53)), 1)
        return MONTH_NAMES[d.month - 1], f'{y}-{MONTH_NAMES[d.month - 1]}', y * 100 + d.month
    except Exception:
        return '', '', 0


def normalize_dataframe(excel_path):
    print(f"엑셀 파일 로드: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name='Raw(1)')
    df.columns = [str(c).strip() for c in df.columns]

    required = ['Year', 'Week', 'Factory', 'Type', 'Category', "Q'ty"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}")

    df = df.dropna(subset=['Year', 'Week', 'Factory', 'Type', 'Category'], how='any').copy()
    df["Q'ty"] = pd.to_numeric(df["Q'ty"], errors='coerce').fillna(0)
    df['Year'] = pd.to_numeric(df['Year'], errors='coerce').fillna(0).astype(int)
    df['Week'] = pd.to_numeric(df['Week'], errors='coerce').fillna(0).astype(int)
    df['Factory'] = df['Factory'].astype(str).str.strip()
    df['Type'] = df['Type'].astype(str).str.strip()
    df['Category'] = df['Category'].astype(str).str.strip()

    if 'Month' not in df.columns or 'YearMonth' not in df.columns or 'Month_Sort' not in df.columns:
        calculated = df.apply(lambda r: week_to_month(r['Year'], r['Week']), axis=1)
        df['Month'] = calculated.apply(lambda x: x[0])
        df['YearMonth'] = calculated.apply(lambda x: x[1])
        df['Month_Sort'] = calculated.apply(lambda x: x[2])
    else:
        # Fill blanks or formula cells not evaluated by pandas
        calculated = df.apply(lambda r: week_to_month(r['Year'], r['Week']), axis=1)
        df['Month'] = df['Month'].fillna('').astype(str)
        df['YearMonth'] = df['YearMonth'].fillna('').astype(str)
        df['Month_Sort'] = pd.to_numeric(df['Month_Sort'], errors='coerce').fillna(0).astype(int)
        mask = (df['Month'].str.strip() == '') | (df['YearMonth'].str.strip() == '') | (df['Month_Sort'] == 0)
        if mask.any():
            df.loc[mask, 'Month'] = calculated[mask].apply(lambda x: x[0])
            df.loc[mask, 'YearMonth'] = calculated[mask].apply(lambda x: x[1])
            df.loc[mask, 'Month_Sort'] = calculated[mask].apply(lambda x: x[2]).astype(int)

    print(f"- 총 행 수: {len(df):,}")
    print(f"- 법인 수: {df['Factory'].nunique()}")
    print(f"- 주차 범위: WK{df['Week'].min():02d} ~ WK{df['Week'].max():02d}")
    print(f"- 월 범위: {df.sort_values('Month_Sort')['YearMonth'].iloc[0]} ~ {df.sort_values('Month_Sort')['YearMonth'].iloc[-1]}")
    return df


def build_period_payload(df_period):
    equipment = {}
    production = {}
    factories = [f for f in FACTORIES_ORDER if f in set(df_period['Factory'])]
    factories += [f for f in sorted(set(df_period['Factory'])) if f not in factories]

    for factory in factories:
        df_fac = df_period[df_period['Factory'] == factory]
        equipment[factory] = {key: get_val(df_fac, 'Machine', cat) for key, cat in EQUIPMENT_MAP.items()}
        production[factory] = {key: get_val(df_fac, 'Performance', cat) for key, cat in PRODUCTION_MAP.items()}

    return {'equipment': equipment, 'production': production}


def convert_weekly(df):
    raw_data = {}
    for year in sorted(df['Year'].unique()):
        year_key = f"{int(year)}Y"
        raw_data[year_key] = {}
        df_year = df[df['Year'] == year]
        for week in sorted(df_year['Week'].unique()):
            week_key = f"WK{int(week):02d}"
            raw_data[year_key][week_key] = build_period_payload(df_year[df_year['Week'] == week])
    return raw_data


def convert_monthly(df):
    monthly_data = {}
    for year in sorted(df['Year'].unique()):
        year_key = f"{int(year)}Y"
        monthly_data[year_key] = {}
        df_year = df[df['Year'] == year].sort_values('Month_Sort')
        months = df_year[['YearMonth', 'Month_Sort']].drop_duplicates().sort_values('Month_Sort')
        for _, row in months.iterrows():
            period_key = str(row['YearMonth'])
            monthly_data[year_key][period_key] = build_period_payload(df_year[df_year['YearMonth'] == period_key])
    return monthly_data


def build_index_html(raw_data, monthly_data, template_path, output_path):
    print(f"HTML 템플릿 로드: {template_path}")
    template = template_path.read_text(encoding='utf-8')
    js_data = (
        "const rawData = " + json.dumps(raw_data, ensure_ascii=False, indent=2) + ";\n" +
        "const monthlyData = " + json.dumps(monthly_data, ensure_ascii=False, indent=2) + ";"
    )
    if '%%RAWDATA_PLACEHOLDER%%' not in template:
        print("❌ 템플릿에 %%RAWDATA_PLACEHOLDER%% 없음")
        sys.exit(1)
    output_html = template.replace('%%RAWDATA_PLACEHOLDER%%', js_data)
    output_path.write_text(output_html, encoding='utf-8')
    print(f"✅ index.html 생성 완료: {len(output_html):,} bytes")


def validate(raw_data, monthly_data):
    latest_year = sorted(raw_data.keys())[-1]
    latest_week = sorted(raw_data[latest_year].keys())[-1]
    latest_month = list(monthly_data[latest_year].keys())[-1]
    print(f"최신 주간: {latest_year} {latest_week}")
    print(f"최신 월간: {latest_year} {latest_month}")


if __name__ == '__main__':
    base_dir = Path(__file__).parent.parent
    excel_path = base_dir / 'Raw Data_Global Equipment-Based Productivity.xlsx'
    template_path = base_dir / 'dashboard_template.html'
    output_path = base_dir / 'index.html'

    if not excel_path.exists():
        print(f"❌ 엑셀 파일 없음: {excel_path}")
        sys.exit(1)
    if not template_path.exists():
        print(f"❌ 템플릿 파일 없음: {template_path}")
        sys.exit(1)

    df = normalize_dataframe(excel_path)
    raw_data = convert_weekly(df)
    monthly_data = convert_monthly(df)
    build_index_html(raw_data, monthly_data, template_path, output_path)
    validate(raw_data, monthly_data)
    print("완료: Weekly / Monthly 전환 대시보드가 생성되었습니다.")
