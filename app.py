import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string
import gspread
from google.oauth2.service_account import Credentials

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 관리 시스템", layout="wide")

# [함수] 글자 정규화
def normalize_text(text):
    if pd.isna(text): return ""
    return str(text).replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")

# [함수] 열 인덱스 안전 변환
def get_col_idx(col_str):
    try:
        return column_index_from_string(col_str.upper()) - 1
    except:
        return None

# [함수] 구글 시트 연동
def append_to_master(df):
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key("1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt")
    worksheet = spreadsheet.sheet1
    worksheet.append_rows(df.values.tolist())
    return True

# URL 설정
DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

st.title("🏗️ 서원건설 - 단가 관리 시스템")

tab1, tab2 = st.tabs(["🏗️ 1. 단가 자동 입력", "➕ 2. 신규 단가 파일 정리"])

# --- [탭 1] ---
with tab1:
    with st.expander("⚙️ [설정] 데이터 위치 옵션", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            start_row = st.number_input("데이터 시작 행", value=4, min_value=1, key="start1")
        with c2:
            col_name = st.text_input("품명 열 (A-Z)", value="A", key="name1")
            col_spec = st.text_input("규격 열 (A-Z)", value="B", key="spec1")
            col_mat = st.text_input("재료비 열 (A-Z)", value="E", key="mat1")
            col_lab = st.text_input("노무비 열 (A-Z)", value="G", key="lab1")
            col_exp = st.text_input("경비 열 (A-Z)", value="I", key="exp1")

    uploaded_file = st.file_uploader("작업할 견적서 엑셀 업로드", type=['xlsx', 'xls'], key="file1")
    
    if uploaded_file:
        wb = openpyxl.load_workbook(uploaded_file)
        ws_name = st.selectbox("작업할 시트 선택", wb.sheetnames, key="sheet1")
        file_ext = st.selectbox("저장할 파일 확장자", [".xlsx", ".xls"], index=0, key="ext1")
        
        if st.button("단가 매칭 실행", type="primary"):
            try:
                # 열 인덱스 검증
                c_n, c_s, c_m, c_l, c_e = get_col_idx(col_name), get_col_idx(col_spec), get_col_idx(col_mat), get_col_idx(col_lab), get_col_idx(col_exp)
                if None in [c_n, c_s, c_m, c_l, c_e]:
                    st.error("알파벳(A, B, C...)을 정확히 입력해주세요.")
                else:
                    raw_df = pd.read_csv(DATA_URL)
                    master_df = raw_df.iloc[:, 0:7].copy()
                    master_df.columns = ['품명', '규격', '단위', '재료비', '노무비', '경비', '비고']
                    master_df = master_df[master_df['품명'] != '품명'].copy()
                    master_df['clean_name'] = master_df['품명'].apply(normalize_text)
                    master_df['clean_spec'] = master_df['규격'].apply(normalize_text)
                    
                    uploaded_file.seek(0)
                    df_ex = pd.read_excel(uploaded_file, sheet_name=ws_name, header=None)
                    ws = wb[ws_name]
                    
                    match_count = 0
                    for i in range(start_row - 1, len(df_ex)):
                        v_n, v_s = normalize_text(df_ex.iloc[i, c_n]), normalize_text(df_ex.iloc[i, c_s])
                        match = master_df[(master_df['clean_name'] == v_n) & (master_df['clean_spec'] == v_s)]
                        
                        if not match.empty:
                            row_idx = i + 1
                            ws[f"{col_mat.upper()}{row_idx}"].value = match['재료비'].values[0]
                            ws[f"{col_lab.upper()}{row_idx}"].value = match['노무비'].values[0]
                            ws[f"{col_exp.upper()}{row_idx}"].value = match['경비'].values[0]
                            match_count += 1
                    
                    output = BytesIO()
                    wb.save(output)
                    st.success(f"🎉 매칭 완료! ({match_count}개)")
                    st.download_button("결과 파일 다운로드", output.getvalue(), file_name=f"매칭결과{file_ext}")
            except Exception as e: st.error(f"오류: {e}")

# --- [탭 2] 신규 단가 데이터 정리기 ---
with tab2:
    st.subheader("➕ 신규 단가 데이터 정리기")
    n1, n2 = st.columns(2)
    with n1:
        n_name = st.text_input("외부 품명 열", value="A", key="name2")
        n_spec = st.text_input("외부 규격 열", value="B", key="spec2")
        n_unit = st.text_input("외부 단위 열", value="C", key="unit2") # 추가됨
    with n2:
        n_mat = st.text_input("외부 재료비 열", value="D", key="mat2")
        n_lab = st.text_input("외부 노무비 열", value="E", key="lab2")
        n_exp = st.text_input("외부 경비 열", value="F", key="exp2")

    new_file = st.file_uploader("지자체 양식 엑셀 업로드", type=['xlsx', 'xls'], key="file2")
    if new_file:
        new_ws_name = st.selectbox("작업할 시트 선택", openpyxl.load_workbook(new_file).sheetnames, key="sheet2")
        if st.button("마스터 형식 변환"):
            c_n, c_s, c_u, c_m, c_l, c_e = get_col_idx(n_name), get_col_idx(n_spec), get_col_idx(n_unit), get_col_idx(n_mat), get_col_idx(n_lab), get_col_idx(n_exp)
            if None in [c_n, c_s, c_u, c_m, c_l, c_e]:
                st.error("열 알파벳을 다시 확인해주세요.")
            else:
                df_raw = pd.read_excel(new_file, sheet_name=new_ws_name, header=None)
                df_result = df_raw.iloc[3:, [c_n, c_s, c_u, c_m, c_l, c_e]].copy()
                df_result.columns = ["품명", "규격", "단위", "재료비", "노무비", "경비"]
                st.dataframe(df_result)
                if st.button("🚀 마스터 시트에 추가"):
                    append_to_master(df_result)
                    st.success("추가 완료!")
