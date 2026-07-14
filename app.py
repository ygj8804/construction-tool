import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string
import gspread
from google.oauth2.service_account import Credentials

# 페이지 설정
st.set_page_config(page_title="서원건설 단가 관리 시스템", layout="wide")

# 글자 정규화
def normalize_text(text):
    if pd.isna(text): return ""
    return str(text).replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")

# 구글 시트 연동
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

# --- [탭 1] 단가 자동 입력 ---
with tab1:
    with st.expander("⚙️ [설정] 데이터 위치 옵션", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            start_row = st.number_input("데이터 시작 행", value=4)
        with col2:
            col_name = st.text_input("품명 열 (A-Z)", value="A")
            col_spec = st.text_input("규격 열 (A-Z)", value="B")
            col_mat = st.text_input("재료비 열 (A-Z)", value="E")
            col_lab = st.text_input("노무비 열 (A-Z)", value="G")
            col_exp = st.text_input("경비 열 (A-Z)", value="I")

    uploaded_file = st.file_uploader("작업할 견적서 엑셀 업로드", type=['xlsx', 'xls'])
    
    if uploaded_file:
        wb = openpyxl.load_workbook(uploaded_file)
        ws_name = st.selectbox("작업할 시트 선택", wb.sheetnames)
        # 파일 업로드 후 확장자 선택창 출력
        file_ext = st.selectbox("저장할 확장자", [".xlsx", ".xls"])
        
        if st.button("단가 매칭 실행"):
            try:
                # 데이터 준비
                raw_df = pd.read_csv(DATA_URL).iloc[:, 0:7]
                raw_df.columns = ['품명', '규격', '단위', '재료비', '노무비', '경비', '비고']
                raw_df['clean_name'] = raw_df['품명'].apply(normalize_text)
                raw_df['clean_spec'] = raw_df['규격'].apply(normalize_text)
                
                ws = wb[ws_name]
                df_ex = pd.read_excel(uploaded_file, sheet_name=ws_name, header=None)
                
                match_count = 0
                c_n = column_index_from_string(col_name.upper()) - 1
                c_s = column_index_from_string(col_spec.upper()) - 1
                
                for i in range(start_row - 1, len(df_ex)):
                    v_n = normalize_text(df_ex.iloc[i, c_n])
                    v_s = normalize_text(df_ex.iloc[i, c_s])
                    
                    match = raw_df[(raw_df['clean_name'] == v_n) & (raw_df['clean_spec'] == v_s)]
                    
                    if not match.empty:
                        row_idx = i + 1
                        # 지정한 열(E, G, I)만 타겟팅
                        targets = [(col_mat, match['재료비'].values[0]), (col_lab, match['노무비'].values[0]), (col_exp, match['경비'].values[0])]
                        
                        for col_char, val in targets:
                            cell = ws[f"{col_char.upper()}{row_idx}"]
                            # 수식이 있으면 skip, 없으면 값 대입
                            if cell.data_type != 'f':
                                cell.value = val
                        match_count += 1
                
                output = BytesIO()
                wb.save(output)
                st.success(f"🎉 단가 매칭 완료! (총 {match_count}개 항목 반영됨)")
                st.download_button("결과 파일 다운로드", output.getvalue(), file_name=f"매칭완료{file_ext}")
            except Exception as e: st.error(f"오류: {e}")

# --- [탭 2] 신규 단가 데이터 정리기 ---
with tab2:
    st.subheader("➕ 신규 단가 데이터 정리기")
    n_mat = st.text_input("외부 재료비 열", value="D")
    n_lab = st.text_input("외부 노무비 열", value="E")
    n_exp = st.text_input("외부 경비 열", value="F")
    
    new_file = st.file_uploader("지자체 양식 업로드", type=['xlsx', 'xls'])
    if new_file:
        if st.button("마스터 시트에 추가"):
            df_new = pd.read_excel(new_file, header=None)
            # 여기는 기존 구조 유지
            st.success("데이터 확인 완료 후 추가됩니다.")
