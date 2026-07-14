import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 관리 시스템", layout="wide")

# [글자 정규화 함수: 공백, 탭, 줄바꿈 제거]
def normalize_text(text):
    if pd.isna(text): return ""
    return str(text).replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")

# 2. 구글 시트 연동
def append_to_master(df):
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key("1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt")
    worksheet = spreadsheet.sheet1
    data_to_add = df.values.tolist()
    worksheet.append_rows(data_to_add)
    return True

DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"
EDIT_URL = "https://docs.google.com/spreadsheets/d/1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt/edit?rtpof=true"

st.title("🏗️ 서원건설 - 단가 관리 시스템")
tab1, tab2 = st.tabs(["🏗️ 1. 단가 자동 입력", "➕ 2. 신규 단가 파일 정리"])

with tab1:
    # (설정 UI 생략 - 이전과 동일)
    start_row = st.number_input("데이터 시작 행", value=4)
    c1, c2 = st.columns(2)
    with c1:
        col_name = st.text_input("품명 열", value="A")
        col_spec = st.text_input("규격 열", value="B")
    with c2:
        col_mat = st.text_input("재료비 열", value="E")
        col_lab = st.text_input("노무비 열", value="G")
        col_exp = st.text_input("경비 열", value="I")

    uploaded_file = st.file_uploader("작업할 견적서 엑셀 업로드", type=['xlsx', 'xls'])
    
    if uploaded_file and st.button("단가 매칭 실행", type="primary"):
        try:
            # 1. 마스터 데이터 로드 (헤더 이름 강제 고정)
            master_df = pd.read_csv(DATA_URL, header=0)
            master_df.columns = ['품명', '규격', '단위', '재료비', '노무비', '경비', '비고']
            
            # 마스터 데이터 미리 공백 제거
            master_df['clean_name'] = master_df['품명'].apply(normalize_text)
            master_df['clean_spec'] = master_df['규격'].apply(normalize_text)
            
            # 2. 엑셀 로드
            uploaded_file.seek(0)
            df_ex = pd.read_excel(uploaded_file, header=None)
            
            wb = openpyxl.load_workbook(uploaded_file)
            ws = wb.active
            
            # 3. 매칭 루프
            for i in range(start_row - 1, len(df_ex)):
                idx_name = column_index_from_string(col_name.upper()) - 1
                idx_spec = column_index_from_string(col_spec.upper()) - 1
                
                # 원본 엑셀 값을 공백 제거
                val_name_clean = normalize_text(df_ex.iloc[i, idx_name])
                val_spec_clean = normalize_text(df_ex.iloc[i, idx_spec])
                
                # 매칭
                match = master_df[(master_df['clean_name'] == val_name_clean) & 
                                  (master_df['clean_spec'] == val_spec_clean)]
                
                if not match.empty:
                    row_idx = i + 1
                    ws[f"{col_mat}{row_idx}"] = match['재료비'].values[0]
                    ws[f"{col_lab}{row_idx}"] = match['노무비'].values[0]
                    ws[f"{col_exp}{row_idx}"] = match['경비'].values[0]
            
            # 4. 결과 저장
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            st.success("🎉 매칭 완료!")
            st.download_button("결과 파일 다운로드", output, f"매칭완료_{uploaded_file.name}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"오류: {e}")

# (탭 2는 기존과 동일하게 유지)
