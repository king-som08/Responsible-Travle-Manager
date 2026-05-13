# ✈️ 여행 예약 가격 비교 도구 (Python/Streamlit)

이 프로젝트는 Python과 Streamlit을 사용하여 구현되었습니다.

## 실행 방법

1. **Python 설치**: Python 3.8 이상의 버전이 설치되어 있어야 합니다.
2. **Tesseract OCR 설치**: 이 앱은 이미지에서 글자를 읽기 위해 Tesseract 엔진을 사용합니다.
   - Windows: [Tesseract installer](https://github.com/UB-Mannheim/tesseract/wiki)에서 설치 후 환경 변수에 경로를 추가하세요.
   - Mac: `brew install tesseract`
   - Linux: `sudo apt-get install tesseract-ocr`
3. **의존성 설치**:
   ```bash
   pip install -r requirements.txt
   ```
4. **앱 실행**:
   ```bash
   streamlit run app.py
   ```

## 주요 기능
- 항공편/숙소 예약 스크린샷 업로드
- Pytesseract를 이용한 자동 가격/정보 추출
- 고정 환율(USD, JPY) 기반 원화 환산
- `st.data_editor`를 통한 데이터 직접 수정 및 정렬
- 엑셀(.xlsx) 파일 내보내기
