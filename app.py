import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")

st.title("🏗️ 서원건설")
st.subheader("공사 내역서 단가 자동 입력기")
st.markdown("---")

# 2. 상단 설정 창 (이곳에서 정한 숫자가 실제 매칭 기준이 됩니다)
with st.expander("⚙️ 엑셀 파일 열 위치 설정 (설정한 열에 맞춰 작업합니다)", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        start_row = st.number_input("데이터 시작 행", value=2, min_value=1)
        col_name = st.number_input("품명(A) 열", value=1, min_value=1)
        col_spec = st.number_input("규격(B) 열", value=2, min_value=1)
    with col2:
        col_unit = st.number_input("단위(C) 열", value=3, min_value=1)
        col_mat = st.number_input("재료비단가(E) 열", value=5, min_value=1)
    with col3:
        col_lab = st.number_input("노무비단가(G) 열", value=7, min_value=1)
        col_exp = st.number_input("경비단가(I) 열", value=9, min_value=1)

# 3. 구글 시트 데이터 로드
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

@st.cache_data(ttl=60) 
def load_master_data():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL, header=None)
        master_data = {}
        for _, row in df.iterrows():
            name = str(row[0]).strip() if pd.notnull(row[0]) else ""
            spec = str(row[1]).strip() if pd.notnull(row[1]) else ""
            key = f"{name}_{spec}"
            master_data[key] = {
                'E': row[4] if len(row) > 4 else None,
                'G': row[6] if len(row) > 6 else None,
                'I': row[8] if len(row) > 8 else None
            }
        return master_data
    except Exception as e:
        return str(e)

master_data = load_master_data()

# 4. 파일 처리
uploaded_file = st.file_uploader("작업할 내역서 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file and isinstance(master_data, dict):
    wb = openpyxl.load_workbook(uploaded_file)
    sheet_names = wb.sheetnames
    selected_sheet = st.selectbox("수정할 시트를 선택하세요", sheet_names)
    ws = wb[selected_sheet]

    if st.button("단가 자동 입력 실행"):
        count = 0
        # 상단 설정창에서 입력한 start_row 부터 시작
        for row in range(start_row, ws.max_row + 1): 
            # 상단 설정창에서 정한 열 번호를 기준으로 값을 읽음
            val_name = ws.cell(row=row, column=col_name).value
            val_spec = ws.cell(row=row, column=col_spec).value
            
            name = str(val_name).strip() if val_name is not None else ""
            spec = str(val_spec).strip() if val_spec is not None else ""
            current_key = f"{name}_{spec}"
            
            if current_key in master_data:
                # 상단 설정창에서 정한 열 번호에 값을 기록
                if master_data[current_key]['E'] is not None:
                    ws.cell(row=row, column=col_mat).value = master_data[current_key]['E']
                if master_data[current_key]['G'] is not None:
                    ws.cell(row=row, column=col_lab).value = master_data[current_key]['G']
                if master_data[current_key]['I'] is not None:
                    ws.cell(row=row, column=col_exp).value = master_data[current_key]['I']
                count += 1
        
        st.success(f"성공! {count}개의 항목이 업데이트되었습니다.")

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button("수정된 파일 다운로드", output, f"수정본_{uploaded_file.name}")

elif uploaded_file and isinstance(master_data, str):
    st.error(f"마스터 데이터 로드 오류: {master_data}")
