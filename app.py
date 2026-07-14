import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")

st.title("🏗️ 서원건설 - 단가 자동 입력기")
st.markdown("---")

# 2. 정갈한 설정창
with st.expander("⚙️ [설정] 데이터 시작 행 및 각 항목 열 위치 설정", expanded=True):
    # 1단 구성: 행 설정
    st.subheader("1. 데이터 시작 행 설정")
    start_row = st.number_input("데이터가 실제로 시작되는 행 번호를 입력하세요 (예: 4)", value=4, min_value=1)
    
    st.markdown("---")
    
    # 2단 구성: 열(Column) 설정
    st.subheader("2. 각 항목별 열(Column) 위치 (알파벳 입력)")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**기본 항목**")
        col_name_str = st.text_input("품명 (열 위치)", value="A")
        col_spec_str = st.text_input("규격 (열 위치)", value="B")
        col_unit_str = st.text_input("단위 (열 위치)", value="C")
        
    with col2:
        st.write("**단가 데이터 입력 위치**")
        col_mat_str = st.text_input("재료비단가 (열 위치)", value="E")
        col_lab_str = st.text_input("노무비단가 (열 위치)", value="G")
        col_exp_str = st.text_input("경비단가 (열 위치)", value="I")

# 3. 마스터 데이터 로드
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

# 4. 파일 업로드 및 기능 실행
st.markdown("---")
uploaded_file = st.file_uploader("작업할 견적서 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file and master_data:
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        selected_sheet = st.selectbox("수정할 시트를 선택하세요", wb.sheetnames)
        ws = wb[selected_sheet]

        if st.button("단가 매칭 실행", type="primary"):
            # 알파벳 -> 숫자 변환
            c_name = column_index_from_string(col_name_str.upper())
            c_spec = column_index_from_string(col_spec_str.upper())
            c_mat = column_index_from_string(col_mat_str.upper())
            c_lab = column_index_from_string(col_lab_str.upper())
            c_exp = column_index_from_string(col_exp_str.upper())

            count = 0
            for row in range(start_row, ws.max_row + 1):
                val_a = str(ws.cell(row=row, column=c_name).value).strip()
                val_b = str(ws.cell(row=row, column=c_spec).value).strip()
                
                key = f"{val_a}_{val_b}"
                
                if key in master_data:
                    ws.cell(row=row, column=c_mat).value = master_data[key]['E']
                    ws.cell(row=row, column=c_lab).value = master_data[key]['G']
                    ws.cell(row=row, column=c_exp).value = master_data[key]['I']
                    count += 1
            
            st.success(f"완료! 총 {count}개의 항목에 단가가 입력되었습니다.")
            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            st.download_button("수정된 파일 다운로드", output, f"매칭완료_{uploaded_file.name}")
            
    except Exception as e:
        st.error(f"오류: 설정한 열 위치(알파벳)가 올바른지 확인해주세요. 상세: {e}")

elif uploaded_file and not master_data:
    st.error("마스터 데이터를 불러오지 못했습니다.")
