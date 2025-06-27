# Product_AutoOrder_Individual_Supplier_v1.0.py
import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import math
import datetime
from typing import Dict, Optional
from pathlib import Path
from io import BytesIO
import plotly.express as px

# --- 1. 기본 설정 및 스타일 (변경 없음) ---
st.set_page_config(page_title="LPI TEAM 자동 납품량 계산 시스템", layout="wide")
st.markdown("""
<style>
.footer { position: fixed; left: 80px; bottom: 20px; font-size: 13px; color: #888; }
.total-cell { width: 100%; text-align: right; font-weight: bold; font-size: 1.1em; padding: 10px 0; }
</style>
""", unsafe_allow_html=True)
st.markdown('<div class="footer">by suhyuk (twodoong@gmail.com)</div>', unsafe_allow_html=True)


# --- 2. 설정 및 상수 정의 (변경 없음) ---
SETTINGS_FILE = 'item_settings.json'
FILE_PATTERN = "현황*.xlsx"
COL_ITEM_CODE = '상품코드'
COL_ITEM_NAME = '상품명'
COL_SPEC = '규격'
COL_BARCODE = '바코드'
COL_UNIT_PRICE = '현구매단가'
COL_SUPPLIER = '매입처'
COL_SALES = '매출수량'
COL_STOCK = '현재고'
EXCLUDE_KEYWORDS = ['배송비', '첫 주문', '쿠폰', '개인결제', '마일리지']
INITIAL_DEFAULT_SETTINGS = {'lead_time': 15, 'safety_stock_rate': 10, 'addition_rate': 0, 'order_unit': 5, 'min_sales': 0}

# --- 3. 핵심 기능 함수 (변경 없음) ---
def load_settings() -> Dict[str, Dict]:
    # 이 함수는 session_state 초기화 로직 변경으로 인해 직접 호출되지는 않게 됩니다.
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            settings = json.load(f)
            if "master_defaults" not in settings:
                settings["master_defaults"] = INITIAL_DEFAULT_SETTINGS.copy()
            else:
                if "min_sales" not in settings["master_defaults"]:
                     settings["master_defaults"]['min_sales'] = INITIAL_DEFAULT_SETTINGS['min_sales']

            for sup_settings in settings.get("defaults", {}).values():
                sup_settings.setdefault('min_sales', settings["master_defaults"]['min_sales'])
            for item_settings in settings.get("overrides", {}).values():
                item_settings.setdefault('min_sales', INITIAL_DEFAULT_SETTINGS['min_sales'])
            return settings
    return {"master_defaults": INITIAL_DEFAULT_SETTINGS.copy(), "defaults": {}, "overrides": {}}

def save_settings(settings: Dict[str, Dict]):
    # Streamlit 클라우드 환경의 읽기 전용 파일 시스템 문제로 이 함수는 호출되지 않도록 수정합니다.
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=4)

def find_latest_file(directory: Path, pattern: str) -> Optional[Path]:
    try:
        files = list(directory.glob(pattern))
        if not files: return None
        return max(files, key=lambda p: p.stat().st_mtime)
    except Exception: return None

def get_min_sales_for_row(row: pd.Series, settings: Dict[str, Dict]) -> int:
    item_code = str(row.get(COL_ITEM_CODE, ''))
    supplier = str(row.get(COL_SUPPLIER, ''))
    master_defaults = settings.get("master_defaults", INITIAL_DEFAULT_SETTINGS)

    if item_code in settings.get("overrides", {}) and 'min_sales' in settings["overrides"][item_code]:
        return settings["overrides"][item_code]['min_sales']
    if supplier in settings.get("defaults", {}) and 'min_sales' in settings["defaults"][supplier]:
        return settings["defaults"][supplier]['min_sales']
    return master_defaults.get('min_sales', 0)

def calculate_order_quantity(df: pd.DataFrame, settings: Dict[str, Dict], period_days: int) -> pd.DataFrame:
    results = []
    master_defaults = settings.get("master_defaults", INITIAL_DEFAULT_SETTINGS)
    default_settings = settings.get("defaults", {})
    override_settings = settings.get("overrides", {})

    for row in df.to_dict('records'):
        item_code = str(row.get(COL_ITEM_CODE, ''))
        supplier = str(row.get(COL_SUPPLIER, ''))
        final_settings = {k: v for k, v in {**master_defaults, **default_settings.get(supplier, {}), **override_settings.get(item_code, {})}.items() if k != 'min_sales'}

        lead_time = final_settings.get('lead_time', 0)
        safety_stock_rate = final_settings.get('safety_stock_rate', 0) / 100
        addition_rate = final_settings.get('addition_rate', 0) / 100
        order_unit = final_settings.get('order_unit', 1)
        if order_unit <= 0: order_unit = 1

        sales_quantity = row.get(COL_SALES, 0)
        current_stock = row.get(COL_STOCK, 0)
        row['추천 납품량'] = 0
        row['초과재고 수량'] = 0

        if period_days > 0:
            avg_daily_sales = sales_quantity / period_days
            sales_during_lead_time = avg_daily_sales * lead_time
            safety_stock = sales_during_lead_time * safety_stock_rate
            reorder_point = sales_during_lead_time + safety_stock
            base_order_quantity = reorder_point - current_stock

            if base_order_quantity <= 0:
                if current_stock > reorder_point * 2 and reorder_point > 0:
                    row['비고'] = "초과재고"
                    row['초과재고 수량'] = current_stock - math.ceil(reorder_point)
                else:
                    row['비고'] = "재고 충분"
            else:
                calculated_quantity = base_order_quantity * (1 + addition_rate)
                final_order_quantity = math.ceil(calculated_quantity / order_unit) * order_unit
                row['추천 납품량'] = int(final_order_quantity)
                if current_stock < final_order_quantity:
                    row['비고'] = "납품 필요 (긴급)"
                else:
                    row['비고'] = "납품 필요"

            row['재고 소진 예상일'] = current_stock / avg_daily_sales if avg_daily_sales > 0 else float('inf')
        else:
            row['비고'] = "기간 1일 이상"
            row['재고 소진 예상일'] = float('inf')

        row['적용된 설정'] = f"L:{lead_time} S:{safety_stock_rate*100:.0f}% A:{addition_rate*100:.0f}% U:{order_unit}"
        results.append(row)
    return pd.DataFrame(results)

def style_remarks(val):
    if val in ['납품 필요 (긴급)', '악성 초과재고']:
        return 'color: #D32F2F; font-weight: bold;'
    return ''

# --- 4. Streamlit UI 구성 ---
title_col1, title_col2 = st.columns([3, 1])
with title_col1:
    st.title("LPI TEAM 자동 납품량 계산 시스템 v1.0")

# ### BUG FIX: st.dialog를 st.expander로 변경하여 버전 호환성 문제 해결 ###
with title_col2:
    with st.expander("📖 시스템 설명"):
        st.markdown("""
        ### 📂 1. 입력 항목 설명
        • **시작일/종료일**: 매출 분석 기간 설정 (기본: 30일)  
        • **제외 매출수량**: 입력값 미만 품목은 계산에서 제외  
        • **리드타임(재발주 기간)(일)**: 납품 후 입고까지 소요 기간(재발주 기간)  
        • **안전재고율(%)**: 리드타임(재발주 기간) 동안 예상 매출의 추가 보유 비율  
        • **가산율(%)**: 계산된 납품량에 추가하는 여유분 비율  
        • **납품단위**: 납품 시 최소 단위 (5개 단위 등)  
        
        ### 📊 2. 긴급 납품 품목 비율 설명
        **■ 안전재고 적용 상세 조건:** • 계산식: (일일 평균 매출 수량 × 리드타임(재발주 기간)) × 안전재고율  
        • 목적: 모자랄 것을 대비하는 추가 여유분  
        • 예시: 일일 20개 판매, 리드타임(재발주 기간) 15일, 안전재고율 10%  
        　→ 기본 추전 납품량 = 20 × 15 = 300개  
        　→ 안전재고 = 300 × 0.1 = 30개 (추가 여유분)  
        　→ 총 추전 납품량 = 300 + 30 = 330개  
        
        **■ 긴급 납품 조건:** • 현재고 < 최종 추천 납품량 (납품량이 클수록 긴급)  
        • 예시: 현재고 250개 < 최종 추천 납품량 350개 → 긴급 납품  
        
        **■ 표시 비율 설정:** • 긴급 납품 품목 중 표시할 상위 비율  
        • 정렬 기준: 추천 납품량이 많은 순서  
        • 예시: 긴급 품목 20개 × 25% = 상위 5개 표시  
        　　　긴급 품목 8개 × 50% = 상위 4개 표시  
        
        ### 🧮 3. 납품 추천 상품 계산 조건
        **■ 계산 공식:** • 일일 평균 매출 수량수량 = 총 매출수량 ÷ 분석기간  
        • 기본 추전 납품량 = 일일 평균 매출 수량 × 리드타임(재발주 기간)  
        • 안전재고 = 기본 추전 납품량 × 안전재고율 (추가 여유분)  
        • 총 추전 납품량 = 기본 추전 납품량 + 안전재고  
        • 기본 납품량 = 총 추전 납품량 - 현재고  
        • 최종 납품량 = 기본 납품량 × (1 + 가산율) → 납품단위로 반올림  
        
        **■ 계산 예시:** • 매출수량: 600개(30일), 현재고: 80개, 리드타임(재발주 기간): 15일, 안전재고율: 10%, 가산율: 5%, 납품단위: 10개  
        • 일일 평균: 600÷30 = 20개  
        • 기본 추전 납품량: 20×15 = 300개  
        • 안전재고: 300×0.1 = 30개 (추가 여유분)  
        • 총 추전 납품량: 300+30 = 330개  
        • 기본 납품량: 330-80 = 250개  
        • 최종 납품량: 250×1.05 = 262.5 → 270개(10개 단위)  
        
        **■ 비고(납품 표시) 판정 기준:** • 납품 필요 (긴급): 현재고 < 최종 추천 납품량  
        • 납품 필요: 기본 납품량 > 0, 현재고 ≥ 최종 추천 납품량  
        • 재고 충분: 기본 납품량 ≤ 0  
        • 초과재고: 현재고 > 총 추전 납품량 × 2  
        
        ### ⚙️ 4. 개별 품목별 설정 설명
        **■ 설정 우선순위:** 1. 개별 품목 설정 (최우선)  
        2. 상품별 전체 기본값  
        
        **■ 사용법 예시:** • 특정 상품(A001)은 리드타임(재발주 기간)이 다른 상품보다 길어서 25일로 설정  
        • 상품별 전체 기본값: 리드타임(재발주 기간) 15일 → 개별 설정: 리드타임(재발주 기간) 25일  
        • 계산 시 A001만 25일 적용, 나머지는 15일 적용  
        
        **■ 실제 적용:** • 납품량 계산 실행 후 상품코드 검색  
        • 개별 설정값 입력 후 저장  
        • 재계산 시 개별 설정값 적용  
        • 기본값 복원으로 개별 설정 삭제 가능  
        
        ### 📦 5. 초과재고 현황 계산 조건
        **■ 초과재고 판정:** 현재고 > 총 추전 발주량 × 2  
        
        **■ 각 컬럼 계산 예시:** • 현재고: 800개, 총 추전 발주량: 330개, 매출수량: 600개(30일), 현구매단가: 1,000원  
        • 초과재고 수량 = 800 - 330 = 470개  
        • 초과재고 비율 = 800 ÷ 600 = 1.3배  
        • 초과재고 금액 = 470 × 1,000 = 470,000원  
        • 재고 소진 예상일 = 800 ÷ 20(일일매출) = 40일  
        
        **■ 악성/일반 구분:** • 전체 초과재고 비율의 중간값을 기준으로 분류  
        • 예시: 중간값이 2.0배인 경우  
        　→ 2.0배 이상: 악성 초과재고 (빨간색 표시)  
        　→ 2.0배 미만: 일반 초과재고  
        """)

    with st.expander("📋 사용 메뉴얼"):
        st.markdown("""
        ### **LPI TEAM 자동 납품량 계산 시스템 - 사용자 메뉴얼 (v1.0)**

        안녕하세요! LPI TEAM 자동 납품량 계산 시스템 사용을 환영합니다.
        이 시스템은 매출현황 데이터와 매입처 제공 설정값을 기반으로 최적의 납품량을 자동 계산합니다.

        ---

        #### **1. 시작 전 준비사항: 필요한 파일들**

        시스템 사용을 위해 **2개의 파일**이 필요합니다:

        **▶ ① 매출현황 파일 (필수)**
        • **파일명**: `현황`으로 시작하는 엑셀 파일 (예: `현황20250626_123028.xlsx`)
        • **위치**: PC의 `다운로드` 폴더 (자동 검색됨)
        • **필수 컬럼**: 상품코드, 상품명, 규격, 바코드, 매출수량, 현구매단가, 현재고, 매입처

        **▶ ② 설정값 파일 (매입처 제공)**
        • **파일명**: `하이온_품목별설정값_YYYYMMDD_HHMMSS.xlsx` 형식
        • **제공처**: 매입처에서 제공받은 설정값 파일
        • **내용**: 납품량 계산을 위한 리드타임, 안전재고율 등의 설정값

        > **✅ 체크포인트**: 두 파일이 모두 준비되었나요? 그럼 시작해보세요!

        ---

        #### **2. 기본 사용 흐름: 3단계 완료!**

        ##### **▶ 1단계: 매출현황 파일 확인**
        1. **[1. 분석 대상 파일 및 기간 설정]** 섹션에서 파일 상태를 확인합니다
        2. **자동 검색**: "✅ 자동으로 찾은 최신 파일" 메시지 확인
        3. **수동 업로드**: 파일이 검색되지 않으면 '수동으로 파일 업로드' 토글 사용
        4. **분석 기간**: 시작일/종료일 설정 (기본 30일)

        ##### **▶ 2단계: 설정값 파일 불러오기**
        1. **[2. 납품 설정 관리]** 섹션을 확장합니다
        2. **'설정 파일을 업로드하세요'** 버튼을 클릭합니다
        3. **매입처 제공 설정값 파일을 선택**합니다
        4. **설정값 확인**: 
           - 마스터 기본값이 파란색 박스에 표시됩니다
           - 품목별 상세 설정이 목록으로 표시됩니다

        ##### **▶ 3단계: 납품량 계산 및 결과 확인**
        1. **🚀 납품량 계산 실행** 버튼을 클릭합니다
        2. **요약 대시보드 확인**: 6개 핵심 지표를 한눈에 파악
           - 추천 품목수, 추천 수량, 예상 금액
           - 초과재고 상품 수, 초과재고 수량, 초과재고 합계
        3. **긴급 납품 그래프**: 가장 시급한 상품들의 시각적 확인
        4. **납품 추천 상품 목록**: 상세한 납품 계획 확인
        5. **📥 엑셀 다운로드**: 결과를 엑셀 파일로 저장

        ---

        #### **3. 주요 기능 상세 설명**

        ##### **📊 요약 대시보드 (6개 지표)**
        • **추천 품목수**: 납품이 필요한 상품의 총 개수
        • **추천 수량**: 모든 납품 추천 상품의 총 수량  
        • **예상 금액**: 추천 수량 기준 예상 납품 비용
        • **초과재고 상품 수**: 재고가 과다한 상품 개수
        • **초과재고 수량**: 과다 재고의 총 수량
        • **초과재고 합계**: 과다 재고의 총 금액

        ##### **🚨 납품 상태 구분**
        • **납품 필요 (긴급)**: 즉시 납품이 필요한 위험 상태 (빨간색)
        • **납품 필요**: 계획된 납품이 필요한 상태
        • **재고 충분**: 당분간 납품이 불필요한 상태
        • **초과재고**: 재고가 과도하여 관리가 필요한 상태

        ##### **⚙️ 설정값 관리**
        • **자동 저장**: 설정값 파일 업로드 시 자동으로 영구 저장
        • **자동 로드**: 프로그램 재시작 시 마지막 설정값 자동 적용
        • **설정 교체**: 새로운 설정값 파일 업로드 시 기존 설정 완전 교체

        ---

        #### **4. 고급 활용 팁**

        ##### **🔄 일상 업무 워크플로우**
        1. **매일**: 프로그램 실행 → 자동으로 저장된 설정값 로드
        2. **주기적**: 최신 매출현황 파일 확인 → 납품량 계산 실행
        3. **설정 변경 시**: 새로운 설정값 파일 업로드 → 자동 저장/적용

        ##### **📈 결과 해석 가이드**
        • **긴급 납품 품목**: 우선순위가 높은 납품 대상
        • **재고 소진 예상일**: 숫자가 작을수록 시급함
        • **초과재고 현황**: 재고 최적화가 필요한 품목들

        ##### **🎯 효율적인 사용법**
        • **정기 점검**: 주 1-2회 정기적인 납품량 계산
        • **긴급 대응**: 예상치 못한 주문 증가 시 즉시 재계산
        • **설정 업데이트**: 매입처에서 새로운 설정값 제공 시 즉시 적용

        ---

        #### **5. 문제 해결 가이드**

        ##### **❓ 자주 묻는 질문**
        • **Q**: 설정값이 표시되지 않아요
        • **A**: 매출현황 파일이 먼저 업로드되어 있는지 확인하세요

        • **Q**: 계산 결과가 이상해요  
        • **A**: 분석 기간과 설정값이 올바른지 확인 후 재계산하세요

        • **Q**: 프로그램을 재시작했는데 설정값이 사라졌어요
        • **A**: 설정값 파일을 다시 업로드하면 자동으로 저장됩니다

        ##### **🔧 해결 단계**
        1. **파일 확인**: 매출현황 파일과 설정값 파일 모두 준비
        2. **순서 준수**: 매출현황 → 설정값 → 계산 실행 순서로 진행  
        3. **재시작**: 문제 발생 시 페이지 새로고침 후 다시 시도

        **더 자세한 도움이 필요하시면 시스템 관리자에게 문의하세요!**
        """)

# --- 이하 모든 코드는 이전과 완전히 동일합니다 ---

# ### 수정 1: Session State 초기화 방식 변경 ###
# 파일에서 설정을 불러오는 대신, 항상 비어있는 기본 설정으로 시작합니다.
if 'settings' not in st.session_state: 
    st.session_state.settings = {"master_defaults": INITIAL_DEFAULT_SETTINGS.copy(), "defaults": {}, "overrides": {}}
    
if 'suppliers' not in st.session_state: st.session_state.suppliers = []
if 'result_df' not in st.session_state: st.session_state.result_df = pd.DataFrame()
if 'searched_item' not in st.session_state: st.session_state.searched_item = None

with st.expander("1. 분석 대상 파일 및 기간 설정", expanded=True):
    # ### 수정 2: 파일 자동 검색 기능 제거 ###
    # info_text_part1 = f"파일 검색 패턴: `{FILE_PATTERN}` (다운로드 폴더에서 찾습니다)" # 이 라인은 혼동을 줄 수 있어 주석 처리
    info_text_part2 = "▶ [상품별 매출 현황] 다운로드 엑셀 파일에는 '상품코드', '상품명', '규격', '바코드', '매출수량', '현구매단가', '현재고', '매입처' 컬럼이 포함되어야 합니다."
    st.markdown(f"<span style='color:blue;'>{info_text_part2}</span>", unsafe_allow_html=True)
    
    target_file_path = None
    
    # --- 스마트 파일 로더 시작 ---
    # 먼저 로컬 PC의 다운로드 폴더가 있는지, 그 안에 파일이 있는지 확인
    downloads_path = Path.home() / "Downloads"
    latest_file = None
    
    # 다운로드 폴더가 실제로 존재할 때만 자동 찾기 시도
    if downloads_path.exists():
        latest_file = find_latest_file(downloads_path, FILE_PATTERN)

    # CASE 1: 로컬 PC에서 파일을 자동으로 찾은 경우
    if latest_file:
        st.success(f"✅ 자동으로 찾은 최신 파일: `{latest_file.name}`")
        # 사용자가 원하면 수동으로 전환할 수 있도록 토글 제공
        manual_upload = st.toggle("수동으로 파일 업로드하기")
        
        if not manual_upload:
            target_file_path = latest_file
        else:
            # 토글을 켜면 수동 업로더 표시
            uploaded_file = st.file_uploader("엑셀 파일을 직접 업로드하세요.", type=['xlsx', 'xls'], key="manual_after_auto")
            if uploaded_file:
                target_file_path = uploaded_file
                
    # CASE 2: 파일을 자동으로 찾지 못한 경우 (웹 서버 환경 또는 PC에 파일이 없는 경우)
    else:
        # 수동 업로드 기능만 깔끔하게 표시
        uploaded_file = st.file_uploader("분석할 현황 엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'], key="manual_only")
        if uploaded_file:
            target_file_path = uploaded_file
    # --- 스마트 파일 로더 끝 ---
    
    st.divider()
    today = datetime.date.today()
    
    date_cols = st.columns(2)
    with date_cols[0]:
        start_date = st.date_input("시작일", value=today - datetime.timedelta(days=30))
    with date_cols[1]:
        end_date = st.date_input("종료일", value=today)

    period_days = 0
    if start_date and end_date and start_date <= end_date:
        period_days = (end_date - start_date).days + 1
        st.info(f"분석 기간은 총 {period_days}일 입니다.")
    else:
        st.error("기간 설정이 올바르지 않습니다.")

if target_file_path:
    try:
        df_for_suppliers = pd.read_excel(target_file_path)
        if COL_SUPPLIER in df_for_suppliers.columns:
            unique_suppliers = sorted([str(s) for s in df_for_suppliers[COL_SUPPLIER].unique() if str(s) != 'nan'])
            st.session_state.suppliers = unique_suppliers
        
        # 현황 파일 데이터를 세션에 저장 (설정값과 매칭용)
        st.session_state.current_data_for_matching = df_for_suppliers
    except Exception:
        st.session_state.suppliers = []
        st.session_state.current_data_for_matching = pd.DataFrame()

with st.expander("2. 납품 설정 관리"):
    with st.container():
        st.markdown("##### [마스터] 상품별 전체 기본값 설정")
        
        # 설정값 불러오기
        uploaded_settings_file = st.file_uploader("설정 파일을 업로드하세요.", type=['xlsx', 'xls'], key="settings_uploader")
        if uploaded_settings_file:
            try:
                # 파일 이름을 session_state에 저장하여 변경 감지
                current_file_name = uploaded_settings_file.name
                
                # 이전 파일명과 비교하여 새로운 파일인지 확인
                if 'last_settings_file' not in st.session_state or st.session_state.last_settings_file != current_file_name:
                    st.session_state.last_settings_file = current_file_name
                    
                    settings_df = pd.read_excel(uploaded_settings_file)
                    
                    # 새로운 설정 파일 업로드 시 기존 설정 완전히 초기화
                    st.session_state.settings["overrides"] = {}
                    st.session_state.loaded_individual_settings = []
                    
                    # 매입처별 기본값 찾기
                    master_row = settings_df[settings_df['설정구분'] == '매입처별 기본값']
                    if not master_row.empty:
                        master_data = master_row.iloc[0]
                        st.session_state.loaded_master_settings = {
                            'lead_time': int(master_data.get('리드타임(재발주기간)(일)', 15)),
                            'safety_stock_rate': int(master_data.get('안전재고율(%)', 10)),
                            'addition_rate': int(master_data.get('가산율(%)', 0)),
                            'order_unit': int(master_data.get('발주단위', 5)),
                            'min_sales': int(master_data.get('제외매출수량', 0))
                        }
                        
                        # 세션 상태의 settings 업데이트
                        st.session_state.settings["master_defaults"] = st.session_state.loaded_master_settings.copy()
                        
                        st.success("설정값이 성공적으로 불러와졌습니다.")
                    
                    # 개별 품목 설정 찾기
                    individual_rows = settings_df[settings_df['설정구분'] == '개별 품목 설정']
                    if not individual_rows.empty:
                        st.session_state.loaded_individual_settings = individual_rows.to_dict('records')
                        
                        # 세션 상태의 overrides 업데이트
                        for setting in st.session_state.loaded_individual_settings:
                            item_code = str(setting.get('상품코드', ''))
                            st.session_state.settings["overrides"][item_code] = {
                                'lead_time': int(setting.get('리드타임(재발주기간)(일)', 0)),
                                'safety_stock_rate': int(setting.get('안전재고율(%)', 0)),
                                'addition_rate': int(setting.get('가산율(%)', 0)),
                                'order_unit': int(setting.get('발주단위', 1)),
                                'min_sales': int(setting.get('제외매출수량', 0))
                            }
                    
                    # 화면 갱신을 위한 rerun (파일이 변경된 경우에만 실행)
                    st.rerun()
                    
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")
        
        # 불러온 마스터 설정값 표시
        if 'loaded_master_settings' in st.session_state:
            master_settings = st.session_state.loaded_master_settings
            st.info(f"리드타임(재발주 기간): {master_settings['lead_time']}일 | 안전재고율: {master_settings['safety_stock_rate']}% | 가산율: {master_settings['addition_rate']}% | 발주단위: {master_settings['order_unit']}개 | 제외 매출수량: {master_settings['min_sales']}개")
        else:
            st.caption("설정값 파일을 불러와 주세요.")
    
    st.divider()
    
    # 품목별 상세 설정 표시
    st.markdown("##### 품목별 상세 설정")
    if 'loaded_individual_settings' in st.session_state:
        individual_settings = st.session_state.loaded_individual_settings
        if individual_settings:
            for i, setting in enumerate(individual_settings, 1):
                item_code = str(setting.get('상품코드', ''))
                lead_time = int(setting.get('리드타임(재발주기간)(일)', 0))
                safety_rate = int(setting.get('안전재고율(%)', 0))
                addition_rate = int(setting.get('가산율(%)', 0))
                order_unit = int(setting.get('발주단위', 0))
                min_sales = int(setting.get('제외매출수량', 0))
                
                # 현황 파일에서 상품 정보 찾기
                item_info = "상품 정보 없음"
                barcode_info = "바코드 없음"
                
                # 현재 업로드된 현황 파일에서 찾기
                if 'current_data_for_matching' in st.session_state and not st.session_state.current_data_for_matching.empty:
                    matching_rows = st.session_state.current_data_for_matching[
                        st.session_state.current_data_for_matching[COL_ITEM_CODE].astype(str) == item_code
                    ]
                    if not matching_rows.empty:
                        row = matching_rows.iloc[0]
                        item_name = str(row.get(COL_ITEM_NAME, ''))
                        spec = str(row.get(COL_SPEC, '')) if COL_SPEC in st.session_state.current_data_for_matching.columns else ''
                        barcode = str(row.get(COL_BARCODE, '')) if COL_BARCODE in st.session_state.current_data_for_matching.columns else ''
                        
                        # 상품명(규격) 형식으로 구성
                        if item_name and item_name != 'nan':
                            item_info = item_name
                            if spec and spec != 'nan' and spec.strip():
                                item_info = f"{item_name} ({spec})"
                        
                        if barcode and barcode != 'nan' and barcode.strip():
                            barcode_info = barcode
                
                st.markdown(f"**{i}. {item_code} ({item_info}), {barcode_info}** | 리드타임(재발주 기간): {lead_time}일 | 안전재고율: {safety_rate}% | 가산율: {addition_rate}% | 발주단위: {order_unit}개 | 제외 매출수량: {min_sales}개")
        else:
            st.caption("개별 품목 설정이 없습니다.")
    else:
        st.caption("설정값 파일을 불러와 주세요.")

st.header("🚀 계산 실행")
if st.button("납품량 계산 실행", type="primary"):
    st.session_state.searched_item = None
    if target_file_path and period_days > 0:
        with st.spinner('데이터를 분석하고 있습니다...'):
            try:
                df = pd.read_excel(target_file_path)
                numeric_cols_to_clean = [COL_UNIT_PRICE, COL_SALES, COL_STOCK]
                for col in numeric_cols_to_clean:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('int64')

                original_item_count = len(df)
                exclude_pattern = '|'.join(EXCLUDE_KEYWORDS)
                df_filtered = df[~df[COL_ITEM_NAME].astype(str).str.contains(exclude_pattern, na=False)].copy()
                keyword_excluded_count = original_item_count - len(df_filtered)

                df_filtered['min_sales_applied'] = df_filtered.apply(get_min_sales_for_row, axis=1, settings=st.session_state.settings)
                df_final_filtered = df_filtered[df_filtered[COL_SALES] >= df_filtered['min_sales_applied']].copy()
                df_final_filtered.drop(columns=['min_sales_applied'], inplace=True)

                sales_excluded_count = len(df_filtered) - len(df_final_filtered)
                st.info(f"총 {original_item_count}개 품목 중, 키워드로 {keyword_excluded_count}개, 매출수량 기준으로 {sales_excluded_count}개를 제외하고 계산합니다.")

                required_cols = [COL_ITEM_CODE, COL_ITEM_NAME, COL_UNIT_PRICE, COL_SUPPLIER, COL_SALES, COL_STOCK]
                if not all(col in df.columns for col in required_cols):
                    missing_cols = [col for col in required_cols if col not in df.columns]
                    st.error(f"엑셀 파일에 필수 컬럼이 없습니다: {', '.join(missing_cols)}")
                else:
                    result_df = calculate_order_quantity(df_final_filtered, st.session_state.settings, period_days)
                    st.session_state.result_df = result_df
                    st.success("납품량 계산이 완료되었습니다.")
            except Exception as e:
                st.error(f"파일 처리 또는 계산 중 오류 발생: {e}")
                st.session_state.result_df = pd.DataFrame()

if not st.session_state.result_df.empty:
    result_df = st.session_state.result_df.copy()
    if COL_SPEC in result_df.columns:
        result_df['상품명 (규격)'] = result_df[COL_ITEM_NAME].astype(str) + result_df[COL_SPEC].apply(lambda x: f' ({x})' if pd.notna(x) and str(x).strip() != '' else '')
    else:
        result_df['상품명 (규격)'] = result_df[COL_ITEM_NAME]
    st.header("📊 요약 대시보드 및 결과 데이터")
    
    df_for_view = result_df
    order_needed_df = df_for_view[df_for_view['추천 납품량'] > 0].copy()
    overstock_df = df_for_view[df_for_view['비고'].isin(['초과재고', '악성 초과재고'])].copy()

    # 요약 대시보드 메트릭 계산
    total_order_items = len(order_needed_df) if not order_needed_df.empty else 0
    total_order_quantity = order_needed_df['추천 납품량'].sum() if not order_needed_df.empty else 0
    
    if not order_needed_df.empty:
        order_needed_df.loc[:, '예상 납품 금액'] = order_needed_df['추천 납품량'] * order_needed_df[COL_UNIT_PRICE]
        total_order_cost = order_needed_df['예상 납품 금액'].sum()
    else:
        total_order_cost = 0
    
    total_overstock_items = len(overstock_df) if not overstock_df.empty else 0
    
    if not overstock_df.empty:
        # 초과재고 비율 계산
        overstock_df.loc[:, '초과재고 비율 (재고/매출)'] = overstock_df[COL_STOCK] / overstock_df[COL_SALES].replace(0, np.nan)
        median_ratio = overstock_df['초과재고 비율 (재고/매출)'].median()
        if pd.notna(median_ratio):
            malignant_rows_mask = overstock_df['초과재고 비율 (재고/매출)'] >= median_ratio
            overstock_df.loc[:, '비고'] = np.where(malignant_rows_mask, "악성 초과재고", "초과재고")
        
        total_overstock_quantity = overstock_df['초과재고 수량'].sum()
        overstock_df.loc[:, '초과재고 금액'] = overstock_df['초과재고 수량'] * overstock_df[COL_UNIT_PRICE]
        total_overstock_cost = overstock_df['초과재고 금액'].sum()
    else:
        total_overstock_quantity = 0
        total_overstock_cost = 0

    # 6개 메트릭 표시
    kpi_cols = st.columns(6)
    kpi_cols[0].metric("추천 품목수", f"{total_order_items} 개")
    kpi_cols[1].metric("추천 수량", f"{total_order_quantity:,.0f} 개")
    kpi_cols[2].metric("예상 금액", f"₩ {total_order_cost:,.0f}")
    kpi_cols[3].metric("초과재고 상품 수", f"{total_overstock_items} 개")
    kpi_cols[4].metric("초과재고 수량", f"{total_overstock_quantity:,.0f} 개")
    kpi_cols[5].metric("초과재고 합계", f"₩ {total_overstock_cost:,.0f}")

    st.divider()
    
    urgent_order_df = df_for_view[df_for_view['비고'] == '납품 필요 (긴급)'].copy()
    if not urgent_order_df.empty:
        display_ratio = st.slider("표시할 긴급 납품 품목 비율 (%)", min_value=10, max_value=100, value=25, step=5)
        num_to_show = math.ceil(len(urgent_order_df) * (display_ratio / 100))
        if num_to_show < 1: num_to_show = 1
        
        graph_data = urgent_order_df.nlargest(num_to_show, '추천 납품량')
        st.subheader(f"긴급 납품 Top {num_to_show}개 (추천량 순)")
        fig = px.bar(graph_data, x='상품명 (규격)', y='추천 납품량', 
                     hover_data=[COL_ITEM_CODE, COL_BARCODE, '현재고', '재고 소진 예상일'],
                     labels={'추천 납품량': '추천 납품 수량', '상품명 (규격)': '상품명'})
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    
    st.header("📑 납품 추천 상품")
    st.caption("추천 납품량이 0보다 큰 품목만 표시됩니다.")
    
    display_columns_order = [
        COL_ITEM_CODE, '상품명 (규격)', COL_BARCODE, COL_STOCK, COL_SALES,
        '재고 소진 예상일', '추천 납품량', '비고', '적용된 설정',
        COL_UNIT_PRICE, '예상 납품 금액'
    ]
    final_display_columns = [col for col in display_columns_order if col in order_needed_df.columns]
    
    if not order_needed_df.empty:
        df_to_display_main = order_needed_df[final_display_columns]
        
        st.dataframe(df_to_display_main.style.format(formatter={
            COL_STOCK: "{:,.0f}", COL_SALES: "{:,.0f}", '추천 납품량': "{:,.0f}",
            COL_UNIT_PRICE: "₩{:,.0f}", '예상 납품 금액': "₩{:,.0f}", '재고 소진 예상일': "{:.0f}"
        }, na_rep='').map(style_remarks, subset=['비고']), use_container_width=True, hide_index=True, height=735)

        st.markdown("<hr style='margin:0.5rem 0; border-top: 2px solid #ccc;'>", unsafe_allow_html=True)
        total_cols = st.columns(len(final_display_columns))
        
        item_count = len(df_to_display_main)
        sum_stock = df_to_display_main[COL_STOCK].sum()
        sum_sales = df_to_display_main[COL_SALES].sum()
        sum_order_qty = df_to_display_main['추천 납품량'].sum()
        sum_order_cost = df_to_display_main.get('예상 납품 금액', pd.Series(0)).sum()
        
        total_cols[0].markdown(f"<div class='total-cell' style='text-align: left;'>합계 ({item_count}개 품목)</div>", unsafe_allow_html=True)
        if COL_STOCK in final_display_columns: total_cols[final_display_columns.index(COL_STOCK)].markdown(f"<div class='total-cell'>{sum_stock:,.0f}</div>", unsafe_allow_html=True)
        if COL_SALES in final_display_columns: total_cols[final_display_columns.index(COL_SALES)].markdown(f"<div class='total-cell'>{sum_sales:,.0f}</div>", unsafe_allow_html=True)
        if '추천 납품량' in final_display_columns: total_cols[final_display_columns.index('추천 납품량')].markdown(f"<div class='total-cell'>{sum_order_qty:,.0f}</div>", unsafe_allow_html=True)
        if '예상 납품 금액' in final_display_columns: total_cols[final_display_columns.index('예상 납품 금액')].markdown(f"<div class='total-cell'>₩ {sum_order_cost:,.0f}</div>", unsafe_allow_html=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 엑셀 다운로드용 컬럼 선택
            excel_columns = [
                COL_ITEM_CODE, '상품명 (규격)', COL_BARCODE, COL_STOCK, COL_SALES,
                '추천 납품량', '비고', '적용된 설정'
            ]
            excel_df = df_to_display_main[excel_columns]
            excel_df.to_excel(writer, index=False, sheet_name='OrderList')
            for column in excel_df:
                column_length = max(excel_df[column].astype(str).map(len).max(), len(column))
                col_idx = excel_df.columns.get_loc(column)
                writer.sheets['OrderList'].set_column(col_idx, col_idx, column_length + 2)
        st.download_button(label="📥 엑셀 다운로드", data=output.getvalue(), file_name=f"납품추천결과_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx")

    st.divider()
    
    st.header("📦 초과재고 현황")
    
    if not overstock_df.empty:
        
        overstock_display_cols_order = [
            COL_ITEM_CODE, '상품명 (규격)', COL_BARCODE, COL_STOCK, '초과재고 수량', COL_SALES, 
            '재고 소진 예상일', '초과재고 비율 (재고/매출)', COL_UNIT_PRICE, '초과재고 금액', '비고'
        ]
        final_overstock_cols = [col for col in overstock_display_cols_order if col in overstock_df.columns]
        df_to_display_overstock = overstock_df[final_overstock_cols]
        
        st.dataframe(df_to_display_overstock.style.format(formatter={
            COL_STOCK: "{:,.0f}", '초과재고 수량': "{:,.0f}", COL_SALES: "{:,.0f}", 
            '재고 소진 예상일': "{:.0f}", '초과재고 비율 (재고/매출)': "{:.1f} 배",
            COL_UNIT_PRICE: "₩{:,.0f}", '초과재고 금액': "₩{:,.0f}"
        }, na_rep='').map(style_remarks, subset=['비고']), use_container_width=True, hide_index=True, height=735)

        st.markdown("<hr style='margin:0.5rem 0; border-top: 2px solid #ccc;'>", unsafe_allow_html=True)
        overstock_total_cols = st.columns(len(final_overstock_cols))
        
        overstock_item_count = len(df_to_display_overstock)
        overstock_sum_stock = df_to_display_overstock[COL_STOCK].sum()
        overstock_sum_over_qty = df_to_display_overstock['초과재고 수량'].sum()
        overstock_sum_sales = df_to_display_overstock[COL_SALES].sum()
        overstock_sum_over_cost = df_to_display_overstock.get('초과재고 금액', pd.Series(0)).sum()
        
        overstock_total_cols[0].markdown(f"<div class='total-cell' style='text-align: left;'>합계 ({overstock_item_count}개 품목)</div>", unsafe_allow_html=True)
        if COL_STOCK in final_overstock_cols: overstock_total_cols[final_overstock_cols.index(COL_STOCK)].markdown(f"<div class='total-cell'>{overstock_sum_stock:,.0f}</div>", unsafe_allow_html=True)
        if '초과재고 수량' in final_overstock_cols: overstock_total_cols[final_overstock_cols.index('초과재고 수량')].markdown(f"<div class='total-cell'>{overstock_sum_over_qty:,.0f}</div>", unsafe_allow_html=True)
        if COL_SALES in final_overstock_cols: overstock_total_cols[final_overstock_cols.index(COL_SALES)].markdown(f"<div class='total-cell'>{overstock_sum_sales:,.0f}</div>", unsafe_allow_html=True)
        if '초과재고 금액' in final_overstock_cols: overstock_total_cols[final_overstock_cols.index('초과재고 금액')].markdown(f"<div class='total-cell'>₩ {overstock_sum_over_cost:,.0f}</div>", unsafe_allow_html=True)

        overstock_output = BytesIO()
        with pd.ExcelWriter(overstock_output, engine='xlsxwriter') as writer:
            df_to_display_overstock.to_excel(writer, index=False, sheet_name='Overstock')
            for column in df_to_display_overstock:
                column_length = max(df_to_display_overstock[column].astype(str).map(len).max(), len(column))
                col_idx = df_to_display_overstock.columns.get_loc(column)
                writer.sheets['Overstock'].set_column(col_idx, col_idx, column_length + 2)
        
        st.download_button(label="📥 초과재고 현황 엑셀 다운로드", data=overstock_output.getvalue(), file_name=f"초과재고현황_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx")
    else:
        st.info("초과재고로 분류된 품목이 없습니다.")