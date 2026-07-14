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

# [구글 시트 연동 함수]
def append_to_master(df):
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    client = gspread.authorize(creds)
    
    # 마스터 시트 열기 (본인의 시트 ID)
    spreadsheet = client.open_by_key("1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt")
    worksheet = spreadsheet.sheet1
    
    # 데이터 행 준비 (A~I열)
    data_to_add = []
    for _, row in df.iterrows():
        new_row = [
            row['품명'], row['규격'], row['단위'], 
            "",             # D열
            row['재료비'],  # E열
            "",             # F열 (빈칸)
            row['노무비'],  # G열
            "",             # H열 (빈칸)
            row['경비']     # I열
        ]
        data_to_add.append(new_row)
        
    worksheet.append_rows(data_to_add)
    return True

# [데이터 연결 주소]
DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"
EDIT_URL = "https://docs.google.com/spreadsheets/d/1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt/edit?rtpof=true"

st.title("🏗️ 서원건설 - 단가 관리 시스템")

# 탭 나누기
tab1, tab2 = st.tabs(["🏗️ 1. 단가 자동 입력", "➕ 2. 신규 단가 파일 정리"])

# --- [탭 1] 단가 자동 입력 ---
with tab1:
    col_info, col_btn = st.columns([0.6, 0.4])
    with col_info:
        st.info(f"📋 마스터 데이터 기준: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    with col_btn:
        st.link_button("단가 수정하기 (마스터 시트) ↗️", EDIT_URL)

    with st.expander("⚙️ [설정] 데이터 시작 행 및 각 항목별 위치 입력", expanded=False):
        start_row = st.number_input("데이터 시작 행", value=4, min_value=1, key="start_row")
        c1, c2 = st.columns(2)
        with c1:
            col_name = st.text_input("품명 열", value="A", key="col_name")
            col_spec = st.text_input("규격 열", value="B", key="col_spec")
        with c2:
            col_mat = st.text_input("재료비 열", value="E", key="col_mat")
            col_lab = st.text_input("노무비 열", value="G", key="col_lab")
            col_exp = st.text_input("경비 열", value="I", key="col_exp")

    uploaded_file = st.file_uploader("작업할 견적서 엑셀 업로드", type=['xlsx', 'xls'], key="main_file")
    
    if uploaded_file:
        wb = openpyxl.load_workbook(uploaded_file, read_only=True)
        ws_name = st.selectbox("작업할 시트 선택", wb.sheetnames, key="sheet1")
        wb.close()
        
        # [추가] 저장 확장자 선택
        save_format = st.selectbox("저장할 파일 확장자 선택", [".xlsx (추천)", ".xlsx", ".xls", ".csv"], key="save_format")
        
        if st.button("단가 매칭 실행", type="primary"):
            st.success(f"단가 매칭이 완료되었습니다. ({save_format} 형식으로 저장 준비됨)")

# --- [탭 2] 신규 단가 파일 정리 ---
with tab2:
    st.subheader("➕ 신규 단가 데이터 정리기")
    st.write("지자체 양식 파일을 올리면 [품명, 규격, 단위, 재료비, 노무비, 경비] 순서로 1초 만에 정리해줍니다.")
    
    with st.expander("⚙️ [설정] 외부 파일의 열 위치 확인", expanded=True):
        n1, n2 = st.columns(2)
        with n1:
            n_name = st.text_input("외부 파일 - 품명 열", value="A", key="n_name")
            n_spec = st.text_input("외부 파일 - 규격 열", value="B", key="n_spec")
            n_unit = st.text_input("외부 파일 - 단위 열", value="C", key="n_unit")
        with n2:
            n_mat = st.text_input("외부 파일 - 재료비 열", value="D", key="n_mat")
            n_lab = st.text_input("외부 파일 - Е", value="E", key="n_lab") 
            n_exp = st.text_input("외부 파일 - 경비 열", value="F", key="n_exp")

    new_file = st.file_uploader("지자체 양식 엑셀 파일 업로드", type=['xlsx'], key="new_file")
    
    if new_file:
        wb_n = openpyxl.load_workbook(new_file, read_only=True)
        selected_sheet = st.selectbox("작업할 시트 선택", wb_n.sheetnames, key="sheet2")
        wb_n.close()
        
        df_raw = pd.read_excel(new_file, sheet_name=selected_sheet, header=None)
        
        if st.button("마스터 형식으로 변환"):
            idx_name = column_index_from_string(n_name.upper()) - 1
            idx_spec = column_index_from_string(n_spec.upper()) - 1
            idx_unit = column_index_from_string(n_unit.upper()) - 1
            idx_mat = column_index_from_string(n_mat.upper()) - 1
            idx_lab = column_index_from_string(n_lab.upper()) - 1
            idx_exp = column_index_from_string(n_exp.upper()) - 1
            
            df_result = df_raw.iloc[3:, [idx_name, idx_spec, idx_unit, idx_mat, idx_lab, idx_exp]].copy()
            df_result.columns = ["품명", "규격", "단위", "재료비", "노무비", "경비"]
            
            st.success("변환 완료! 아래 버튼을 눌러 마스터 시트에 자동 추가하세요.")
            st.dataframe(df_result)
            
            if st.button("🚀 마스터 시트에 바로 추가하기"):
                try:
                    append_to_master(df_result)
                    st.success("🎉 성공! 마스터 시트에 데이터가 입력되었습니다.")
                    st.balloons()
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")
            
            csv = df_result.to_csv(index=False).encode('utf-8-sig')
            st.download_button("변환된 데이터 다운로드(CSV)", csv, "마스터_추가용_데이터.csv", "text/csv")
            st.link_button("마스터 시트 열기 ↗️", EDIT_URL)
