import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string
from datetime import datetime

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")

# [중요] 수정이 필요한 주소들
# DATA_URL: 프로그램이 데이터를 읽어오는 전용 주소 (기존 주소 그대로 사용)
DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"
# EDIT_URL: 유 대리님이 수정할 때 쓰는 웹 주소 (브라우저 주소창 주소를 여기에 복사하세요)
EDIT_URL = "https://docs.google.com/spreadsheets/d/본인의_스프레드시트_ID_주소를_여기에_넣으세요"

# 제목 및 상단 정보
st.title("🏗️ 서원건설 - 단가 자동 입력기")

col_info, col_btn = st.columns([0.6, 0.4])
with col_info:
    st.info(f"📋 마스터 데이터 기준: {datetime.now().strftime('%Y-%m-%d')}")
with col_btn:
    # 수정: 다운로드 링크가 아니라 수정용 링크로 변경
    st.link_button("단가 수정하기 ↗️", EDIT_URL)

st.markdown("---")

# 2. 정갈한 설정창
with st.expander("⚙️ [설정] 데이터 시작 행 및 각 항목별 위치 (열) 입력", expanded=True):
    st.subheader("1. 데이터 시작 행")
    start_row = st.number_input("데이터가 실제로 시작되는 행 번호", value=4, min_value=1)
    
    st.markdown("---")
    
    st.subheader("2. 각 항목별 위치 (열) 입력 (알파벳 입력)")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**기본 항목**")
        col_name_str = st.text_input("품명 (열)", value="A")
        col_spec_str = st.text_input("규격 (열)", value="B")
        col_unit_str = st.text_input("단위 (열)", value="C")
        
    with col2:
        st.write("**단가 데이터 입력 위치**")
        col_mat_str = st.text_input("재료비단가 (열)", value="E")
        col_lab_str = st.text_input("노무비단가 (열)", value="G")
        col_exp_str = st.text_input("경비단가 (열)", value="I")

# 3. 마스터 데이터 로드
@st.cache_data(ttl=60)
def load_master_data():
    try:
        df = pd.read_csv(DATA_URL, header=None)
        master_data = {}
        for _, row in df.iterrows():
            key = f"{str(row[0]).strip()}_{str(row[1]).strip()}"
            master_data[key] = {'E': row[4], 'G': row[6], 'I': row[8]}
        return master_data
    except: return None

master_data = load_master_data()

# 4. 파일 업로드 및 기능 실행
st.markdown("---")
uploaded_file = st.file_uploader("작업할 견적서 엑셀 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])
file_ext = st.selectbox("저장할 파일 형식 선택", ["xlsx", "csv"])

if uploaded_file and master_data:
    try:
        if uploaded_file.name.endswith('.csv'):
            st.warning("CSV 파일은 셀 서식이 유지되지 않을 수 있습니다.")
        else:
            wb = openpyxl.load_workbook(uploaded_file)
            ws = wb[wb.sheetnames[0]]

        if st.button("단가 매칭 실행", type="primary"):
            c_name = column_index_from_string(col_name_str.upper())
            c_spec = column_index_from_string(col_spec_str.upper())
            c_mat = column_index_from_string(col_mat_str.upper())
            c_lab = column_index_from_string(col_lab_str.upper())
            c_exp = column_index_from_string(col_exp_str.upper())

            if uploaded_file.name.endswith('.csv'):
                st.error("CSV 직접 수정은 서식 파괴 위험이 있어 .xlsx 파일 사용을 권장합니다.")
            else:
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
                
                final_name = f"매칭완료_{uploaded_file.name.split('.')[0]}.{file_ext}"
                st.download_button(label=f"수정된 파일 다운로드 ({file_ext})", data=output, file_name=final_name)
            
    except Exception as e:
        st.error(f"오류: 설정한 열 위치(알파벳)가 올바른지 확인해주세요. 상세: {e}")

elif uploaded_file and not master_data:
    st.error("마스터 데이터를 불러오지 못했습니다.")
