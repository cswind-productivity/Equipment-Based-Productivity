#!/usr/bin/env python3
"""
CS Wind Dashboard - Excel to HTML Auto Converter
GitHub Actions에서 자동 실행되는 스크립트
Raw Data_Power BI.xlsx → index.html 자동 생성
"""

import pandas as pd
import json
import sys
import os
from pathlib import Path

def get_val(df_fac, type_, cat):
    """특정 Type/Category의 합계 반환, NaN은 0으로 처리"""
    v = df_fac[(df_fac['Type'] == type_) & (df_fac['Category'] == cat)]["Q'ty"].sum()
    return float(v) if not pd.isna(v) else 0.0

def convert_excel_to_rawdata(excel_path):
    """엑셀 파일을 JS rawData 구조로 변환"""
    print(f"📂 엑셀 파일 로드: {excel_path}")
    
    df = pd.read_excel(excel_path, sheet_name='Raw(1)')
    df["Q'ty"] = pd.to_numeric(df["Q'ty"], errors='coerce').fillna(0)
    
    print(f"  - 총 행 수: {len(df)}")
    print(f"  - 법인 수: {df['Factory'].nunique()}")
    print(f"  - 주차 범위: WK{df['Week'].min():02d} ~ WK{df['Week'].max():02d}")

    raw_data = {}

    for year in sorted(df['Year'].unique()):
        year_key = f"{int(year)}Y"
        raw_data[year_key] = {}
        df_year = df[df['Year'] == year]

        for week in sorted(df_year['Week'].unique()):
            week_key = f"WK{int(week):02d}"
            df_week = df_year[df_year['Week'] == week]

            equipment = {}
            production = {}

            for factory in df_week['Factory'].unique():
                df_fac = df_week[df_week['Factory'] == factory]
                equipment[factory] = {
                    "bending":    get_val(df_fac, 'Machine', 'Roll Bending Machine'),
                    "lw":         get_val(df_fac, 'Machine', 'L/W Machine'),
                    "cw":         get_val(df_fac, 'Machine', 'C/W Machine'),
                    "growing":    get_val(df_fac, 'Machine', 'Growing Line'),
                    "paintBooth": get_val(df_fac, 'Machine', 'Paint Booth'),
                    "paintLine":  get_val(df_fac, 'Machine', 'Paint Line'),
                }
                production[factory] = {
                    "bending": get_val(df_fac, 'Performance', 'Bending'),
                    "lw":      get_val(df_fac, 'Performance', 'L/W'),
                    "cw":      get_val(df_fac, 'Performance', 'C/W'),
                    "btgt":    get_val(df_fac, 'Performance', 'BT GT'),
                    "wtgt":    get_val(df_fac, 'Performance', 'WT GT'),
                }

            raw_data[year_key][week_key] = {
                "equipment": equipment,
                "production": production
            }

    return raw_data

def build_index_html(raw_data, template_path, output_path):
    """템플릿 HTML에 rawData 삽입하여 index.html 생성"""
    print(f"📄 HTML 템플릿 로드: {template_path}")
    
    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()

    js_data = "const rawData = " + json.dumps(raw_data, ensure_ascii=False, indent=2) + ";"
    
    if '%%RAWDATA_PLACEHOLDER%%' not in template:
        print("❌ 템플릿에 플레이스홀더 없음!")
        sys.exit(1)

    output_html = template.replace('%%RAWDATA_PLACEHOLDER%%', js_data)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(output_html)

    print(f"✅ index.html 생성 완료: {len(output_html):,} bytes ({len(output_html)/1024:.1f} KB)")

def validate(raw_data):
    """최신 주차 데이터 검증 출력"""
    year = list(raw_data.keys())[0]
    weeks = sorted(raw_data[year].keys())
    latest_week = weeks[-1]
    
    # 실제 데이터가 있는 최신 주차 찾기
    for w in reversed(weeks):
        has_data = any(
            raw_data[year][w]['production'][f]['bending'] > 0
            for f in raw_data[year][w]['production']
        )
        if has_data:
            latest_week = w
            break

    print(f"\n📊 최신 실적 주차: {latest_week}")
    prod = raw_data[year][latest_week]['production']
    equip = raw_data[year][latest_week]['equipment']
    
    factories_order = ['VN #1', 'VN #2', 'TW', 'CN', 'TR #1', 'TR #2', 'AM', 'PT On', 'PT Off']
    print(f"  {'법인':<10} {'장비':>4}  {'생산':>5}  {'효율':>7}")
    print(f"  {'-'*35}")
    for fac in factories_order:
        if fac in prod:
            eq = equip[fac]['bending']
            pr = prod[fac]['bending']
            eff = pr/eq if eq > 0 else 0
            print(f"  {fac:<10} {eq:>4.0f}대  {pr:>5.0f}  {eff:>6.2f}")

if __name__ == "__main__":
    # 경로 설정
    base_dir = Path(__file__).parent.parent  # 저장소 루트
    excel_path = base_dir / "data" / "Raw Data_Power BI.xlsx"
    template_path = base_dir / "dashboard_template.html"
    output_path = base_dir / "index.html"

    # 파일 존재 확인
    if not excel_path.exists():
        print(f"❌ 엑셀 파일 없음: {excel_path}")
        sys.exit(1)
    if not template_path.exists():
        print(f"❌ 템플릿 파일 없음: {template_path}")
        sys.exit(1)

    # 변환 실행
    raw_data = convert_excel_to_rawdata(excel_path)
    build_index_html(raw_data, template_path, output_path)
    validate(raw_data)
    
    print("\n🎉 완료! index.html이 최신 데이터로 업데이트되었습니다.")
