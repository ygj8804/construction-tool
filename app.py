import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")

st.title("🏗️ 서원건설 - 단가 자동 입력기")
st.markdown("---")

# 2. 필수 설정창 (알파벳 입력 방식)
with st.expander("⚙️ [필수] 엑셀 열 위치 설정 (A, B, C, E, G, I 등 알파벳으로 입력하세요)", expanded=True):
    st.write("견적서 엑셀 파일의 각 항목이 몇 번째 열인지 알파벳(A, B, C...)으로 입력하세요.")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        start_row = st.number_input("데이터 시작 행 (이미지상 4행)", value=4, min_value=1)
        st.markdown("---")
        col_name_str = st.text_input("품명 (A열) 위치", value="A")
        col_spec_str = st.text_input("규격 (B열) 위치", value="B")
        col_unit_str = st.text_input("단위 (C열) 위치", value="C")
        
    with col2:
        st.write("### 단가 데이터가 들어갈 위치")
        col_mat_str = st.text_input("재료비단가 (E열) 위치", value="E")
        col_lab_str = st.text_input("노무비단가 (G열) 위치", value="G")
        col_exp_str = st.text_input("경비단가 (I열) 위치", value="I")

# 3. 마스터 데이터 로드 (구글 시트)
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

@st.cache_data(ttl=60)
def load_master_data():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL, header=None)
        master_data = {}
        for _, row in df.iterrows():
            key = f"{str(row[0]).strip()}_{str(row[1]).strip()}"
            master_data[key] = {
                'E': row[4], 'G': row[6], 'I': row[8]
            }
        return master_data
    except: return None

master_data = load_master_data()

# 4. 파일 업로드 및 기능 실행
uploaded_file = st.file_uploader("작업할 견적서 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file and master_data:
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        selected_sheet = st.selectbox("수정할 시트를 선택하세요", wb.sheetnames)
        ws = wb[selected_sheet]

        if st.button("단가 매칭 실행"):
            # 알파벳을 숫자 인덱스로 변환 (A->1, B->2, E->5 등)
            col_name = column_index_from_string(col_name_str.upper())
            col_spec = column_index_from_string(col_spec_str.upper())
            col_mat = column_index_from_string(col_mat_str.upper())
            col_lab = column_index_from_string(col_lab_str.upper())
            col_exp = column_index_from_string(col_exp_str.upper())

            count = 0
            for row in range(start_row, ws.max_row + 1):
                # 변환된 숫자 인덱스를 사용하여 셀 값 가져오기
                val_a = str(ws.cell(row=row, column=col_name).value).strip()
                val_b = str(ws.cell(row=row, column=col_spec).value).strip()
                
                key = f"{val_a}_{val_b}"
                
                if key in master_data:
                    ws.cell(row=row, column=col_mat).value = master_data[key]['E']
                    ws.cell(row=row, column=col_lab).value = master_data[key]['G']
                    ws.cell(row=row, column=col_exp).value = master_data[key]['I']
                    count += 1
            
            st.success(f"완료! {count}개의 항목에 단가가 입력되었습니다.")
            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            st.download_button("수정된 파일 다운로드", output, f"매칭완료_{uploaded_file.name}")
            
    except Exception as e:
        st.error(f"오류 발생: 알파벳(A, B, E 등)을 정확히 입력했는지 확인해주세요. 상세: {e}")

elif uploaded_file and not master_data:
    st.error("마스터 데이터를 불러오지 못했습니다.")
