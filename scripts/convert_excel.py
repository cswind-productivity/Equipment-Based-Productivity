#!/usr/bin/env python3
"""
CS Wind Dashboard - Excel to HTML Auto Converter
Weekly + Monthly rawData 생성 버전
"""
import pandas as pd
import json
import sys
from pathlib import Path


def get_val(df_fac, type_, cat):
    """특정 Type/Category의 합계 반환, NaN은 0으로 처리"""
    v = df_fac[(df_fac['Type'] == type_) & (df_fac['Category'] == cat)]["Q'ty"].sum()
    return float(v) if not pd.isna(v) else 0.0


def add_month_columns(df):
    """GitHub Actions에서 Excel 수식값을 읽지 못하는 경우를 대비하여 Python에서 월간 컬럼 재계산"""
    df = df.copy()
    df['Year'] = pd.to_numeric(df['Year'], errors='coerce').fillna(0).astype(int)
    df['Week'] = pd.to_numeric(df['Week'], errors='coerce').fillna(0).astype(int)
    base_date = pd.to_datetime(df['Year'].astype(str) + '-01-01', errors='coerce') + pd.to_timedelta((df['Week'] - 1) * 7, unit='D')
    df['Month'] = base_date.dt.strftime('%b')
    df['YearMonth'] = base_date.dt.strftime('%Y-%b')
    df['Month_Sort'] = base_date.dt.year * 100 + base_date.dt.month
    return df


def convert_by_period(df, period_col, period_key_func=None):
    """Week 또는 YearMonth 기준 rawData 구조 생성"""
    raw_data = {}
    for year in sorted(df['Year'].dropna().unique()):
        year_key = f"{int(year)}Y"
        raw_data[year_key] = {}
        df_year = df[df['Year'] == year]

        if period_col == 'Week':
            periods = sorted(df_year['Week'].dropna().unique())
        else:
            periods = [x for _, x in sorted(zip(df_year['Month_Sort'], df_year['YearMonth']))]
            periods = list(dict.fromkeys(periods))

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
                    "bending": get_val(df_fac, 'Machine', 'Roll Bending Machine'),
                    "lw": get_val(df_fac, 'Machine', 'L/W Machine'),
                    "cw": get_val(df_fac, 'Machine', 'C/W Machine'),
                    "growing": get_val(df_fac, 'Machine', 'Growing Line'),
                    "paintBooth": get_val(df_fac, 'Machine', 'Paint Booth'),
                    "paintLine": get_val(df_fac, 'Machine', 'Paint Line'),
                }
                production[factory] = {
                    "bending": get_val(df_fac, 'Performance', 'Bending'),
                    "lw": get_val(df_fac, 'Performance', 'L/W'),
                    "cw": get_val(df_fac, 'Performance', 'C/W'),
                    "btgt": get_val(df_fac, 'Performance', 'BT GT'),
                    "wtgt": get_val(df_fac, 'Performance', 'WT GT'),
                }
            raw_data[year_key][period_key] = {
                "equipment": equipment,
                "production": production
            }
    return raw_data


def convert_excel_to_rawdata(excel_path):
    print(f" 엑셀 파일 로드: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name='Raw(1)')
    df["Q'ty"] = pd.to_numeric(df["Q'ty"], errors='coerce').fillna(0)
    df = add_month_columns(df)
    print(f" - 총 행 수: {len(df)}")
    print(f" - 법인 수: {df['Factory'].nunique()}")
    print(f" - 주차 범위: WK{df['Week'].min():02d} ~ WK{df['Week'].max():02d}")
    print(f" - 월 범위: {df.sort_values('Month_Sort')['YearMonth'].iloc[0]} ~ {df.sort_values('Month_Sort')['YearMonth'].iloc[-1]}")
    weekly_raw_data = convert_by_period(df, 'Week')
    monthly_raw_data = convert_by_period(df, 'YearMonth')
    return weekly_raw_data, monthly_raw_data


def build_index_html(weekly_raw_data, monthly_raw_data, template_path, output_path):
    print(f" HTML 템플릿 로드: {template_path}")
    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()
    js_data = "const rawData = " + json.dumps(weekly_raw_data, ensure_ascii=False, indent=2) + ";\n"
    js_data += "const monthlyData = " + json.dumps(monthly_raw_data, ensure_ascii=False, indent=2) + ";"
    if '%%RAWDATA_PLACEHOLDER%%' not in template:
        print("❌ 템플릿에 플레이스홀더 없음!")
        sys.exit(1)
    output_html = template.replace('%%RAWDATA_PLACEHOLDER%%', js_data)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(output_html)
    print(f"✅ index.html 생성 완료: {len(output_html):,} bytes ({len(output_html)/1024:.1f} KB)")


if __name__ == "__main__":
    base_dir = Path(__file__).parent.parent
    excel_path = base_dir / "Raw Data_Global Equipment-Based Productivity.xlsx"
    template_path = base_dir / "dashboard_template.html"
    output_path = base_dir / "index.html"

    if not excel_path.exists():
        print(f"❌ 엑셀 파일 없음: {excel_path}")
        sys.exit(1)
    if not template_path.exists():
        print(f"❌ 템플릿 파일 없음: {template_path}")
        sys.exit(1)

    weekly_raw_data, monthly_raw_data = convert_excel_to_rawdata(excel_path)
    build_index_html(weekly_raw_data, monthly_raw_data, template_path, output_path)
    print("\n 완료! Weekly + Monthly 데이터가 index.html에 반영되었습니다.")
