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

# 2. 구글 시트 연동 함수
def append_to_master(df):
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key("1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt")
    worksheet = spreadsheet.sheet1
    
    data_to_add = []
    for _, row in df.iterrows():
        new_row = [row['품명'], row['규격'], row['단위'], "", row['재료비'], "", row['노무비'], "", row['경비']]
        data_to_add.append(new_row)
    worksheet.append_rows(data_to_add)
    return True

# 3. URL 설정
EDIT_URL = "https://docs.google.com/spreadsheets/d/1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt/edit?rtpof=true"
DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

st.title("🏗️ 서원건설 - 단가 관리 시스템")

tab1, tab2 = st.tabs(["🏗️ 1. 단가 자동 입력", "➕ 2. 신규 단가 파일 정리"])

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
        wb = openpyxl.load_workbook(uploaded_file)
        ws_name = st.selectbox("작업할 시트 선택", wb.sheetnames, key="sheet1")
        ws = wb[ws_name]
        
        if st.button("단가 매칭 실행", type="primary"):
            try:
                # 데이터 불러오기
                master_df = pd.read_csv(DATA_URL)
                
                # [중요 체크] 마스터 시트에 필요한 열이 있는지 확인
                required_cols = ['품명', '규격', '재료비', '노무비', '경비']
                missing_cols = [c for c in required_cols if c not in master_df.columns]
                
                if missing_cols:
                    st.error(f"오류: 마스터 시트에 다음 열이 없습니다: {missing_cols}")
                    st.write("현재 마스터 시트에서 인식한 열 목록:", master_df.columns.tolist())
                    st.write("구글 시트의 첫 번째 줄(헤더)에 정확히 해당 이름들이 있는지 확인해 주세요.")
                else:
                    # 엑셀 파일 안전하게 읽기
                    uploaded_file.seek(0)
                    df_ex = pd.read_excel(uploaded_file, sheet_name=ws_name, header=None)
                    
                    for i in range(start_row - 1, len(df_ex)):
                        row_idx = i + 1
                        idx_name = column_index_from_string(col_name.upper()) - 1
                        idx_spec = column_index_from_string(col_spec.upper()) - 1
                        
                        val_name = str(df_ex.iloc[i, idx_name])
                        val_spec = str(df_ex.iloc[i, idx_spec])
                        
                        match = master_df[(master_df['품명'].astype(str) == val_name) & (master_df['규격'].astype(str) == val_spec)]
                        
                        if not match.empty:
                            ws[f"{col_mat}{row_idx}"] = match['재료비'].values[0]
                            ws[f"{col_lab}{row_idx}"] = match['노무비'].values[0]
                            ws[f"{col_exp}{row_idx}"] = match['경비'].values[0]
                    
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.success("🎉 단가 매칭이 완료되었습니다!")
                    st.download_button("결과 파일 다운로드 (서식 유지)", output, f"매칭완료_{uploaded_file.name}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    wb.close()
            except Exception as e:
                st.error(f"예기치 않은 오류 발생: {e}")

# --- [탭 2] ---
with tab2:
    st.subheader("➕ 신규 단가 데이터 정리기")
    # (탭 2 내용은 기존과 동일하므로 생략하거나 유지하시면 됩니다)
    with st.expander("⚙️ [설정] 외부 파일의 열 위치 확인", expanded=True):
        n1, n2 = st.columns(2)
        with n1:
            n_name = st.text_input("품명 열", value="A", key="n_name")
            n_spec = st.text_input("규격 열", value="B", key="n_spec")
            n_unit = st.text_input("단위 열", value="C", key="n_unit")
        with n2:
            n_mat = st.text_input("재료비 열", value="D", key="n_mat")
            n_lab = st.text_input("노무비 열", value="E", key="n_lab")
            n_exp = st.text_input("경비 열", value="F", key="n_exp")

    new_file = st.file_uploader("지자체 양식 엑셀 파일 업로드", type=['xlsx'], key="new_file")
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
