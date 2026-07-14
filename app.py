import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")

st.title("🏗️ 서원건설")
st.subheader("공사 내역서 단가 자동 입력기")
st.markdown("---")

# 2. 설정 창 (유 대리님 요청대로 A, B, C, E, G, I 순서대로 배치)
with st.expander("⚙️ 내역서 양식 설정 (열 번호 입력)", expanded=True):
    st.write("사용하시는 엑셀 파일의 열 번호를 입력하세요 (A=1, B=2, C=3, E=5, G=7, I=9 등)")
    
    start_row = st.number_input("데이터 시작 행", value=2, min_value=1)
    
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: col_a = st.number_input("품명(A) 열", value=1, min_value=1)
    with c2: col_b = st.number_input("규격(B) 열", value=2, min_value=1)
    with c3: col_c = st.number_input("단위(C) 열", value=3, min_value=1)
    with c4: col_e = st.number_input("재료비(E) 열", value=5, min_value=1)
    with c5: col_g = st.number_input("노무비(G) 열", value=7, min_value=1)
    with c6: col_i = st.number_input("경비(I) 열", value=9, min_value=1)

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
    selected_sheet = st.selectbox("수정할 시트를 선택하세요", wb.sheetnames)
    ws = wb[selected_sheet]

    if st.button("단가 자동 입력 실행"):
        count = 0
        for row in range(start_row, ws.max_row + 1): 
            # 설정창의 번호를 기준으로 값을 가져옴
            val_a = ws.cell(row=row, column=col_a).value
            val_b = ws.cell(row=row, column=col_b).value
            
            name = str(val_a).strip() if val_a is not None else ""
            spec = str(val_b).strip() if val_b is not None else ""
            current_key = f"{name}_{spec}"
            
            if current_key in master_data:
                # 설정창의 번호(E, G, I)에 값을 입력
                if master_data[current_key]['E'] is not None:
                    ws.cell(row=row, column=col_e).value = master_data[current_key]['E']
                if master_data[current_key]['G'] is not None:
                    ws.cell(row=row, column=col_g).value = master_data[current_key]['G']
                if master_data[current_key]['I'] is not None:
                    ws.cell(row=row, column=col_i).value = master_data[current_key]['I']
                count += 1
        
        st.success(f"성공! {count}개의 항목이 업데이트되었습니다.")

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button("수정된 파일 다운로드", output, f"수정본_{uploaded_file.name}")

elif uploaded_file and isinstance(master_data, str):
    st.error(f"마스터 데이터 로드 오류: {master_data}")
