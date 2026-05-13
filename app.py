# ✈️ 여행 예약 가격 비교 도구 - Requirements
# 아래 라이브러리를 설치해야 정상적으로 실행됩니다.
# pip install streamlit pandas pytesseract Pillow opencv-python xlsxwriter

import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image
import cv2
import numpy as np
import re
import io
import xlsxwriter
from datetime import datetime

# ==========================================
# 1. 기본 설정 및 안내문
# ==========================================
st.set_page_config(page_title="✈️ 여행 예약 가격 비교 도구", layout="wide")

# OCR 사용을 위한 Tesseract 경로 설정 (로컬 환경에 맞춰 수정 필요할 수 있음)
# 예: pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

st.title("✈️ 여행 예약 가격 비교 도구")
st.warning("⚠️ 항공편과 숙소의 가격은 매일 변동될 수 있으므로, 정확한 비교를 위해 가급적 당일 캡처한 정보를 한 번에 입력할 것을 권장합니다.")

# ==========================================
# 2. 세션 상태(Session State) 관리
# ==========================================
if 'df' not in st.session_state:
    # 컬럼 정의: ["사이트명", "구분", "상세정보", "결제방식", "원본가격(단위)", "환산가격(KRW)"]
    st.session_state.df = pd.DataFrame(columns=["사이트명", "구분", "상세정보", "결제방식", "원본가격(단위)", "환산가격(KRW)"])

# ==========================================
# 3. 데이터 입력 및 이미지 업로드 UI
# ==========================================
with st.container():
    st.subheader("📝 새로운 정보 추가")
    col1, col2 = st.columns(2)
    
    with col1:
        site_name = st.text_input("사이트명 (예: 인터파크, 아고다, 스카이스캐너)", placeholder="사이트 이름을 입력하세요")
    
    with col2:
        category = st.selectbox("구분", ["항공편", "숙소"])

    # 이미지 업로드 섹션
    uploaded_images = []
    if category == "항공편":
        sub_col1, sub_col2 = st.columns(2)
        with sub_col1:
            img1 = st.file_uploader("1번: 기본 정보 사진 (항공사, 시간 등)", type=["jpg", "jpeg", "png"])
            if img1:
                st.image(img1, width=150, caption="기본 정보")
                uploaded_images.append(img1)
        with sub_col2:
            img2 = st.file_uploader("2번: 최종 결제 사진 (가격 정보 포함)", type=["jpg", "jpeg", "png"])
            if img2:
                st.image(img2, width=150, caption="결제 정보")
                uploaded_images.append(img2)
    else:  # 숙소
        img_hostel = st.file_uploader("통합 정보 사진 (숙소명, 가격 정보 포함)", type=["jpg", "jpeg", "png"])
        if img_hostel:
            st.image(img_hostel, width=150, caption="숙소 정보")
            uploaded_images.append(img_hostel)

    # 버튼 활성화 조건
    images_ready = (category == "항공편" and len(uploaded_images) == 2) or (category == "숙소" and len(uploaded_images) == 1)
    submit_ready = site_name.strip() != "" and images_ready

# ==========================================
# 4. OCR 파싱 및 환율 변환 로직
# ==========================================
def perform_ocr(images):
    """
    업로드된 이미지들에서 텍스트를 추출하고 정보를 파싱합니다.
    """
    full_text = ""
    for img_file in images:
        # PIL 이미지로 변환
        image = Image.open(img_file)
        # OCR 수행 (한글+영어 인지 권장)
        try:
            text = pytesseract.image_to_string(image, lang='kor+eng')
            full_text += "\n" + text
        except Exception as e:
            st.error(f"OCR 처리 중 오류 발생: {e}. Tesseract가 설치되어 있는지 확인하세요.")
            return None

    # --- 1) 가격 및 통화 추출 (Regex) ---
    # 숫자와 통화 기호 조합 찾기 (예: 45,000, 1,200.50, ₩50000, $100 등)
    # ₩, ¥, $, KRW, JPY, USD 대응
    price_pattern = re.compile(r'([₩¥\$]|(?:KRW|JPY|USD))\s*([\d,.]+)|([\d,.]+)\s*([₩¥\$]|(?:KRW|JPY|USD))', re.IGNORECASE)
    price_match = price_pattern.search(full_text)
    
    raw_price_str = "(수동 입력 필요)"
    converted_price = 0.0
    
    # 환율 설정 (예시 임의 고정값)
    EXCHANGE_RATES = {
        'JPY': 9.0,   # 100엔당 900원이면 1엔당 9원
        '¥': 9.0,
        'USD': 1350.0,
        '$': 1350.0,
        'KRW': 1.0,
        '₩': 1.0
    }

    if price_match:
        # 그룹 매칭 결과 정리
        g = price_match.groups()
        currency = g[0] or g[3]
        value_str = g[1] or g[2]
        
        if currency and value_str:
            clean_value = float(value_str.replace(',', ''))
            raw_price_str = f"{currency.upper()}{clean_value:,.0f}"
            
            # 환산 로직
            cur_upper = currency.upper()
            rate = EXCHANGE_RATES.get(cur_upper, 1.0)
            
            # 엔화(%¥)인 경우 보통 100 단위로 계산되는 경우가 많으나 여기선 1단위 환율 적용
            converted_price = clean_value * rate
            if cur_upper in ['JPY', '¥']:
                # 만약 입력값이 100엔 단위가 아니라 1엔 단위라고 가정하면 그대로 곱함
                pass 
    else:
        # 가격 기호가 없을 경우 숫자만이라도 찾아봄
        simple_price = re.search(r'[\d,]{3,}', full_text)
        if simple_price:
            val = float(simple_price.group().replace(',', ''))
            raw_price_str = f"₩{val:,.0f}"
            converted_price = val

    # --- 2) 상세 정보 추출 (항공사, 시간, 숙소명 등) ---
    # 항공사 추출 (간단 예시)
    airlines = ["대한항공", "아시아나", "제주항공", "진에어", "티웨이", "에어부산", "Korean Air", "Asiana"]
    detail_info = "(수동 입력 필요)"
    for a in airlines:
        if a.lower() in full_text.lower():
            detail_info = a
            break
    
    # 숙소명 추출 (간단 예시: 'Hotel' 또는 '호텔' 근처 단어)
    if category == "숙소":
        hotel_match = re.search(r'([가-힣\w\s]+(?:호텔|Hotel|Resort|리조트|Stay|스테이))', full_text)
        if hotel_match:
            detail_info = hotel_match.group(1).strip()

    # --- 3) 결제 방식 추출 ---
    payment_methods = ["신용카드", "계좌이체", "네이버페이", "카카오페이", "현장결제", "선결제"]
    pay_method = "(수동 입력 필요)"
    for p in payment_methods:
        if p in full_text:
            pay_method = p
            break

    return {
        "상세정보": detail_info,
        "결제방식": pay_method,
        "원본가격(단위)": raw_price_str,
        "환산가격(KRW)": int(converted_price)
    }

# 목록 추가 버튼 처리
if st.button("목록에 추가 (OCR 분석)", disabled=not submit_ready):
    with st.spinner("이미지를 분석하고 있습니다..."):
        result = perform_ocr(uploaded_images)
        if result:
            new_row = {
                "사이트명": site_name,
                "구분": category,
                "상세정보": result["상세정보"],
                "결제방식": result["결제방식"],
                "원본가격(단위)": result["원본가격(단위)"],
                "환산가격(KRW)": result["환산가격(KRW)"]
            }
            st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([new_row])], ignore_index=True)
            st.success("데이터가 성공적으로 추출되어 추가되었습니다!")

# ==========================================
# 5. 결과 테이블 및 정렬/수정 기능
# ==========================================
st.divider()
st.subheader("📊 가격 비교 리스트")
st.info("💡 표의 셀을 더블클릭하여 내용을 직접 수정할 수 있습니다. 환산가격 컬럼을 클릭하면 정렬됩니다.")

if not st.session_state.df.empty:
    # st.data_editor를 사용하여 실시간 수정 가능하게 구현
    # num_rows="dynamic"은 행 추가/삭제 가능
    edited_df = st.data_editor(
        st.session_state.df,
        use_container_width=True,
        num_rows="fixed", # 요구사항에 맞춰 행 추가는 상단 폼에서만
        column_config={
            "환산가격(KRW)": st.column_config.NumberColumn(
                "환산가격(KRW)",
                format="%d ₩",
                help="임의 환율이 적용된 원화 가격입니다."
            ),
            "구분": st.column_config.SelectboxColumn(
                "구분",
                options=["항공편", "숙소"],
                required=True
            )
        },
        key="data_editor"
    )
    
    # 수정된 내용 세션에 반영
    if not edited_df.equals(st.session_state.df):
        st.session_state.df = edited_df

    # ==========================================
    # 6. 엑셀 파일 내보내기 (Export)
    # ==========================================
    def to_excel(df):
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Price_Comparison')
        
        # 엑셀 서식 지정 (선택사항)
        workbook = writer.book
        worksheet = writer.sheets['Price_Comparison']
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        writer.close()
        processed_data = output.getvalue()
        return processed_data

    excel_data = to_excel(st.session_state.df)
    
    st.download_button(
        label="📥 결과 엑셀 파일로 다운로드",
        data=excel_data,
        file_name=f"travel_price_comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("전체 데이터 초기화"):
        st.session_state.df = pd.DataFrame(columns=["사이트명", "구분", "상세정보", "결제방식", "원본가격(단위)", "환산가격(KRW)"])
        st.rerun()

else:
    st.write("아직 등록된 정보가 없습니다. 상단에서 정보를 입력해주세요.")

# 하단 도움말
with st.expander("ℹ️ OCR 인식율을 높이는 팁"):
    st.write("""
    1. **글자가 선명하게** 보이도록 캡처하세요.
    2. **다크 모드**보다는 라이트 모드 화면이 인식이 잘 됩니다.
    3. 가격 앞뒤에 **통화 단위(KRW, ₩, $, USD 등)**가 포함되도록 캡처하세요.
    4. 인식에 실패할 경우, 표에서 직접 수정이 가능합니다.
    """)
