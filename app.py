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

# 탭 구성
tab1, tab2 = st.tabs(["🏗️ 1. 단가 자동 입력", "➕ 2. 신규 단가 파일 정리"])

# --- [탭 1] 단가 자동 입력 ---
with tab1:
    with st.expander("⚙️ [설정] 데이터 위치 옵션", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            start_row = st.number_input("데이터 시작 행", value=4, min_value=1, key="start1")
        with c2:
            col_name = st.text_input("품명 열 (알파벳)", value="A", key="name1")
            col_spec = st.text_input("규격 열 (알파벳)", value="B", key="spec1")
            col_mat = st.text_input("재료비 열 (알파벳)", value="E", key="mat1")
            col_lab = st.text_input("노무비 열 (알파벳)", value="G", key="lab1")
            col_exp = st.text_input("경비 열 (알파벳)", value="I", key="exp1")

    uploaded_file = st.file_uploader("작업할 견적서 엑셀 업로드", type=['xlsx', 'xls'], key="file1")
    
    if uploaded_file:
        wb = openpyxl.load_workbook(uploaded_file)
        ws_name = st.selectbox("작업할 시트 선택", wb.sheetnames, key="sheet1")
        file_ext = st.selectbox("저장할 파일 확장자", [".xlsx", ".xls"], index=0, key="ext1")
        
        if st.button("단가 매칭 실행", type="primary"):
            try:
                # 마스터 데이터 로드 (7개 열만 추출)
                raw_df = pd.read_csv(DATA_URL)
                master_df = raw_df.iloc[:, 0:7].copy() 
                master_df.columns = ['품명', '규격', '단위', '재료비', '노무비', '경비', '비고']
                master_df['clean_name'] = master_df['품명'].apply(normalize_text)
                master_df['clean_spec'] = master_df['규격'].apply(normalize_text)
                
                uploaded_file.seek(0)
                df_ex = pd.read_excel(uploaded_file, sheet_name=ws_name, header=None)
                ws = wb[ws_name]
                
                match_count = 0
                
                for i in range(start_row - 1, len(df_ex)):
                    idx_name = column_index_from_string(col_name.upper()) - 1
                    idx_spec = column_index_from_string(col_spec.upper()) - 1
                    
                    val_name_clean = normalize_text(df_ex.iloc[i, idx_name])
                    val_spec_clean = normalize_text(df_ex.iloc[i, idx_spec])
                    
                    match = master_df[(master_df['clean_name'] == val_name_clean) & 
                                      (master_df['clean_spec'] == val_spec_clean)]
                    
                    if not match.empty:
                        row_idx = i + 1
                        # [핵심] 수식 보호 로직: 셀에 수식('f')이 있다면 덮어쓰지 않음
                        for col_char in [col_mat, col_lab, col_exp]:
                            cell = ws[f"{col_char}{row_idx}"]
                            if cell.data_type != 'f': # 수식이 아닌 경우에만 값 입력
                                if col_char == col_mat: cell.value = match['재료비'].values[0]
                                if col_char == col_lab: cell.value = match['노무비'].values[0]
                                if col_char == col_exp: cell.value = match['경비'].values[0]
                        match_count += 1
                
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"🎉 단가 매칭 완료! (총 {match_count}개 항목 매칭됨)")
                st.download_button(
                    label=f"결과 파일 다운로드 {file_ext}",
                    data=output,
                    file_name=f"매칭완료_{uploaded_file.name.split('.')[0]}{file_ext}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                wb.close()
            except Exception as e:
                st.error(f"오류 발생: {e}")

# --- [탭 2] 신규 단가 데이터 정리기 ---
with tab2:
    st.subheader("➕ 신규 단가 데이터 정리기")
    n1, n2 = st.columns(2)
    with n1:
        n_name = st.text_input("외부 품명 열", value="A", key="name2")
        n_spec = st.text_input("외부 규격 열", value="B", key="spec2")
    with n2:
        n_mat = st.text_input("외부 재료비 열", value="D", key="mat2")
        n_lab = st.text_input("외부 노무비 열", value="E", key="lab2")
        n_exp = st.text_input("외부 경비 열", value="F", key="exp2")

    new_file = st.file_uploader("지자체 양식 엑셀 업로드", type=['xlsx', 'xls'], key="file2")
    
    if new_file:
        wb_new = openpyxl.load_workbook(new_file)
        new_ws_name = st.selectbox("작업할 시트 선택", wb_new.sheetnames, key="sheet2")
        
        if st.button("마스터 형식 변환"):
            df_raw = pd.read_excel(new_file, sheet_name=new_ws_name, header=None)
            # 3행(데이터 시작 행)부터 처리
            df_result = df_raw.iloc[3:, [
                column_index_from_string(n_name.upper()) - 1,
                column_index_from_string(n_spec.upper()) - 1,
                0, column_index_from_string(n_mat.upper()) - 1,
                column_index_from_string(n_lab.upper()) - 1,
                column_index_from_string(n_exp.upper()) - 1
            ]].copy()
            df_result.columns = ["품명", "규격", "단위", "재료비", "노무비", "경비"]
            st.dataframe(df_result)
            
            if st.button("🚀 마스터 시트에 추가"):
                try:
                    append_to_master(df_result)
                    st.success("성공! 마스터 시트에 추가되었습니다.")
                except Exception as e:
                    st.error(f"오류: {e}")

# [푸터]
st.markdown("---")
st.caption("개발: 유강진")
