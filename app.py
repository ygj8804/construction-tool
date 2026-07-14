import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")

st.title("🏗️ 서원건설 - 단가 자동 입력기")
st.markdown("---")

# 2. 필수 설정창 (이 설정을 통해 현장 양식에 맞게 위치를 지정합니다)
with st.expander("⚙️ [필수] 엑셀 열 번호 및 행 설정 (현장 양식에 맞춰 숫자를 수정하세요)", expanded=True):
    st.write("사용하시는 견적서 파일의 각 항목이 몇 번째 열(Column)에 있는지 적어주세요.")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        start_row = st.number_input("데이터 시작 행 (보통 2~3)", value=2, min_value=1)
        st.markdown("---")
        col_name = st.number_input("품명 (A열) 위치", value=1, min_value=1)
        col_spec = st.number_input("규격 (B열) 위치", value=2, min_value=1)
        col_unit = st.number_input("단위 (C열) 위치", value=3, min_value=1)
        
    with col2:
        st.write("### 단가 데이터가 들어갈 위치")
        col_mat = st.number_input("재료비단가 (E열) 위치", value=5, min_value=1)
        col_lab = st.number_input("노무비단가 (G열) 위치", value=7, min_value=1)
        col_exp = st.number_input("경비단가 (I열) 위치", value=9, min_value=1)

# 3. 마스터 데이터 로드 (구글 시트)
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

@st.cache_data(ttl=60)
def load_master_data():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL, header=None)
        master_data = {}
        for _, row in df.iterrows():
            # A열(0) + B열(1)을 키로 사용
            key = f"{str(row[0]).strip()}_{str(row[1]).strip()}"
            master_data[key] = {
                'E': row[4], # 재료비
                'G': row[6], # 노무비
                'I': row[8]  # 경비
            }
        return master_data
    except: return None

master_data = load_master_data()

# 4. 파일 업로드 및 기능 실행
uploaded_file = st.file_uploader("작업할 견적서 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file and master_data:
    wb = openpyxl.load_workbook(uploaded_file)
    
    # [기능 유지] 시트 선택 기능
    selected_sheet = st.selectbox("수정할 시트를 선택하세요", wb.sheetnames)
    ws = wb[selected_sheet]

    if st.button("단가 매칭 실행"):
        count = 0
        # [기능 유지] 설정된 행부터 끝까지 반복
        for row in range(start_row, ws.max_row + 1):
            val_a = str(ws.cell(row=row, column=col_name).value).strip()
            val_b = str(ws.cell(row=row, column=col_spec).value).strip()
            
            key = f"{val_a}_{val_b}"
            
            if key in master_data:
                # [기능 유지] 설정한 열 위치에 값 입력
                ws.cell(row=row, column=col_mat).value = master_data[key]['E']
                ws.cell(row=row, column=col_lab).value = master_data[key]['G']
                ws.cell(row=row, column=col_exp).value = master_data[key]['I']
                count += 1
        
        st.success(f"완료! {count}개의 항목에 단가가 입력되었습니다.")
        
        # [기능 유지] 파일 다운로드 기능
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button("수정된 파일 다운로드", output, f"매칭완료_{uploaded_file.name}")
elif uploaded_file and not master_data:
    st.error("마스터 데이터를 불러오지 못했습니다. 구글 시트 연결을 확인하세요.")
