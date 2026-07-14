import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# 1. 페이지 설정 및 초기화
st.set_page_config(page_title="서원건설 단가 관리 시스템", layout="wide")

# [핵심] 공백 제거 함수 (매칭 오류 방지)
def normalize_text(text):
    if pd.isna(text): return ""
    return str(text).replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")

# 2. 구글 시트 연동 함수
def append_to_master(df):
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key("1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt")
    worksheet = spreadsheet.sheet1
    data_to_add = df.values.tolist()
    worksheet.append_rows(data_to_add)
    return True

# 3. URL 설정
EDIT_URL = "https://docs.google.com/spreadsheets/d/1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt/edit?rtpof=true"
DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

st.title("🏗️ 서원건설 - 단가 관리 시스템")

# 4. 탭 정의
tab1, tab2 = st.tabs(["🏗️ 1. 단가 자동 입력", "➕ 2. 신규 단가 파일 정리"])

# --- [탭 1] 단가 자동 입력 ---
with tab1:
    col_info, col_btn = st.columns([0.6, 0.4])
    with col_info:
        st.info(f"📋 마스터 데이터 기준: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    with col_btn:
        st.link_button("단가 수정하기 (마스터 시트) ↗️", EDIT_URL)

    with st.expander("⚙️ [설정] 데이터 위치 및 시작 행", expanded=False):
        start_row = st.number_input("데이터 시작 행", value=4, min_value=1, key="start_row")
        c1, c2 = st.columns(2)
        with c1:
            col_name = st.text_input("품명 열", value="A")
            col_spec = st.text_input("규격 열", value="B")
        with c2:
            col_mat = st.text_input("재료비 열", value="E")
            col_lab = st.text_input("노무비 열", value="G")
            col_exp = st.text_input("경비 열", value="I")

    # 파일 업로더 (확장자 제한)
    uploaded_file = st.file_uploader("작업할 견적서 엑셀 업로드", type=['xlsx', 'xls'], key="file1")
    
    if uploaded_file:
        wb = openpyxl.load_workbook(uploaded_file)
        ws_name = st.selectbox("작업할 시트 선택", wb.sheetnames, key="sheet_sel")
        ws = wb[ws_name]
        
        if st.button("단가 매칭 실행", type="primary"):
            try:
                # 마스터 데이터 로드 (헤더 강제 고정)
                master_df = pd.read_csv(DATA_URL, header=0)
                master_df.columns = ['품명', '규격', '단위', '재료비', '노무비', '경비', '비고']
                
                # 마스터 데이터 공백 제거 (매칭용)
                master_df['clean_name'] = master_df['품명'].apply(normalize_text)
                master_df['clean_spec'] = master_df['규격'].apply(normalize_text)
                
                # 엑셀 데이터 읽기
                uploaded_file.seek(0)
                df_ex = pd.read_excel(uploaded_file, sheet_name=ws_name, header=None)
                
                # 매칭 루프
                for i in range(start_row - 1, len(df_ex)):
                    idx_name = column_index_from_string(col_name.upper()) - 1
                    idx_spec = column_index_from_string(col_spec.upper()) - 1
                    
                    val_name_clean = normalize_text(df_ex.iloc[i, idx_name])
                    val_spec_clean = normalize_text(df_ex.iloc[i, idx_spec])
                    
                    match = master_df[(master_df['clean_name'] == val_name_clean) & 
                                      (master_df['clean_spec'] == val_spec_clean)]
                    
                    if not match.empty:
                        row_idx = i + 1
                        ws[f"{col_mat}{row_idx}"] = match['재료비'].values[0]
                        ws[f"{col_lab}{row_idx}"] = match['노무비'].values[0]
                        ws[f"{col_exp}{row_idx}"] = match['경비'].values[0]
                
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                st.success("🎉 단가 매칭이 완료되었습니다!")
                st.download_button("결과 파일 다운로드", output, f"매칭완료_{uploaded_file.name}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                wb.close()
            except Exception as e:
                st.error(f"오류: {e}")

# --- [탭 2] 신규 단가 파일 정리 ---
with tab2:
    st.subheader("➕ 신규 단가 데이터 정리기")
    with st.expander("⚙️ [설정] 외부 파일 열 위치", expanded=True):
        n1, n2 = st.columns(2)
        with n1:
            n_name = st.text_input("품명 열", value="A", key="n_name")
            n_spec = st.text_input("규격 열", value="B", key="n_spec")
            n_unit = st.text_input("단위 열", value="C", key="n_unit")
        with n2:
            n_mat = st.text_input("재료비 열", value="D", key="n_mat")
            n_lab = st.text_input("노무비 열", value="E", key="n_lab")
            n_exp = st.text_input("경비 열", value="F", key="n_exp")

    new_file = st.file_uploader("지자체 양식 엑셀 파일 업로드", type=['xlsx'], key="file2")
    if new_file:
        df_raw = pd.read_excel(new_file, header=None)
        if st.button("마스터 형식으로 변환"):
            idx_name = column_index_from_string(n_name.upper()) - 1
            idx_spec = column_index_from_string(n_spec.upper()) - 1
            idx_unit = column_index_from_string(n_unit.upper()) - 1
            idx_mat = column_index_from_string(n_mat.upper()) - 1
            idx_lab = column_index_from_string(n_lab.upper()) - 1
            idx_exp = column_index_from_string(n_exp.upper()) - 1
            
            df_result = df_raw.iloc[3:, [idx_name, idx_spec, idx_unit, idx_mat, idx_lab, idx_exp]].copy()
            df_result.columns = ["품명", "규격", "단위", "재료비", "노무비", "경비"]
            
            st.dataframe(df_result)
            if st.button("🚀 마스터 시트에 바로 추가하기"):
                try:
                    append_to_master(df_result)
                    st.success("🎉 성공! 마스터 시트에 데이터가 입력되었습니다.")
                except Exception as e:
                    st.error(f"오류: {e}")
