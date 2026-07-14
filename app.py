import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="서원건설 단가 입력기", layout="wide")

st.title("🏗️ 서원건설 - 단가 자동 입력기")

# 1. 설정 창: 여기서 입력하는 숫자가 실제 데이터의 열 위치(Column Index)입니다.
with st.expander("⚙️ [필수] 엑셀 열 위치 설정 (현장 양식에 맞게 숫자를 수정하세요)", expanded=True):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        start_row = st.number_input("데이터가 시작되는 행 번호", value=2)
        st.markdown("---")
        col_name = st.number_input("품명(A) 위치(열 번호)", value=1)
        col_spec = st.number_input("규격(B) 위치(열 번호)", value=2)
        col_unit = st.number_input("단위(C) 위치(열 번호)", value=3)
    
    with col2:
        st.write("### 단가 데이터 위치")
        col_mat = st.number_input("재료비단가(E) 열 번호", value=5)
        col_lab = st.number_input("노무비단가(G) 열 번호", value=7)
        col_exp = st.number_input("경비단가(I) 열 번호", value=9)

# 2. 데이터 로드 (고정)
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

@st.cache_data(ttl=60)
def load_master_data():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL, header=None)
        master_data = {}
        for _, row in df.iterrows():
            key = f"{str(row[0]).strip()}_{str(row[1]).strip()}"
            master_data[key] = {'E': row[4], 'G': row[6], 'I': row[8]}
        return master_data
    except: return None

master_data = load_master_data()

# 3. 파일 처리
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=['xlsx'])

if uploaded_file and master_data:
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb[wb.sheetnames[0]]

    if st.button("단가 매칭 실행"):
        count = 0
        for row in range(start_row, ws.max_row + 1):
            # 설정창에서 정한 열 번호(col_name, col_spec)를 그대로 사용
            val_a = str(ws.cell(row=row, column=col_name).value).strip()
            val_b = str(ws.cell(row=row, column=col_spec).value).strip()
            
            key = f"{val_a}_{val_b}"
            
            if key in master_data:
                # 설정창에서 정한 열 번호(col_mat, col_lab, col_exp)에 값을 덮어씀
                ws.cell(row=row, column=col_mat).value = master_data[key]['E']
                ws.cell(row=row, column=col_lab).value = master_data[key]['G']
                ws.cell(row=row, column=col_exp).value = master_data[key]['I']
                count += 1
        
        st.success(f"완료! {count}개 항목 매칭됨")
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button("다운로드", output, "수정완료.xlsx")
