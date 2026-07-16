import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string
from datetime import datetime
from openpyxl.styles import Alignment

# [기존 설정 유지]
st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")
DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"
EDIT_URL = "https://docs.google.com/spreadsheets/d/1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt/edit?rtpof=true"

st.title("🏗️ 서원건설 - 단가 자동 입력기")

# 데이터 정제 함수 (공백 및 특수문자 제거)
def clean_str(text):
    if pd.isna(text): return ""
    # 모든 공백(일반 공백 + 특수 공백) 제거 및 문자열화
    return str(text).replace('\xa0', ' ').replace('\u200b', '').strip()

@st.cache_data(ttl=60)
def load_master_data(duplicate_option):
    try:
        df = pd.read_csv(DATA_URL, header=None)
        temp_data = {} 
        for _, row in df.iterrows():
            # 키 생성 시에도 정제 적용
            key = f"{clean_str(row[0])}_{clean_str(row[1])}"
            if key not in temp_data: temp_data[key] = {'E': [], 'G': [], 'I': []}
            try:
                temp_data[key]['E'].append(float(str(row[4]).replace(',', '')))
                temp_data[key]['G'].append(float(str(row[6]).replace(',', '')))
                temp_data[key]['I'].append(float(str(row[8]).replace(',', '')))
            except: continue
        
        master_data = {}
        for key, values in temp_data.items():
            if not values['E']: continue
            if duplicate_option == "최저가 적용":
                master_data[key] = {'E': min(values['E']), 'G': min(values['G']), 'I': min(values['I'])}
            else:
                master_data[key] = {'E': max(values['E']), 'G': max(values['G']), 'I': max(values['I'])}
        return master_data
    except Exception as e: 
        st.error(f"데이터 로드 실패: {e}")
        return None

# [UI 레이아웃 유지]
col_info, col_btn = st.columns([0.6, 0.4])
with st.expander("⚙️ [설정] 데이터 시작 행 및 각 항목별 위치 (열) 입력", expanded=True):
    start_row = st.number_input("데이터 시작 행", value=4, min_value=1)
    duplicate_option = st.selectbox("중복 데이터 처리", ["최저가 적용", "최고가 적용"])
    col1, col2 = st.columns(2)
    with col1:
        col_name_str = st.text_input("품명 (열)", value="A")
        col_spec_str = st.text_input("규격 (열)", value="B")
    with col2:
        col_mat_str = st.text_input("재료비단가 (열)", value="E")
        col_lab_str = st.text_input("노무비단가 (열)", value="G")
        col_exp_str = st.text_input("경비단가 (열)", value="I")

master_data = load_master_data(duplicate_option)
with col_info:
    st.info(f"📋 기준 시간: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    if master_data: st.metric("총 마스터 데이터 항목", f"{len(master_data):,} 개")
with col_btn: st.link_button("단가 수정하기 ↗️", EDIT_URL)

uploaded_file = st.file_uploader("엑셀 파일 업로드", type=['xlsx', 'xls', 'csv'])
file_ext = st.selectbox("저장 형식", ["xlsx", "csv"])

if uploaded_file and master_data:
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        selected_sheet = st.selectbox("작업할 시트", wb.sheetnames)
        ws = wb[selected_sheet]

        if st.button("단가 매칭 실행", type="primary"):
            c_name = column_index_from_string(col_name_str.strip().upper())
            c_spec = column_index_from_string(col_spec_str.strip().upper())
            c_mat = column_index_from_string(col_mat_str.strip().upper())
            c_lab = column_index_from_string(col_lab_str.strip().upper())
            c_exp = column_index_from_string(col_exp_str.strip().upper())

            count, unmatched = 0, []
            for row in range(start_row, ws.max_row + 1):
                val_a = clean_str(ws.cell(row=row, column=c_name).value)
                val_b = clean_str(ws.cell(row=row, column=c_spec).value)
                key = f"{val_a}_{val_b}"
                
                if key in master_data:
                    ws.cell(row=row, column=c_mat).value = master_data[key]['E']
                    ws.cell(row=row, column=c_lab).value = master_data[key]['G']
                    ws.cell(row=row, column=c_exp).value = master_data[key]['I']
                    count += 1
                elif val_a or val_b: # 값이 있는데 매칭 안 된 경우
                    unmatched.append(f"{val_a} / {val_b}")
            
            st.success(f"완료! {count}개 항목 입력.")
            if unmatched:
                st.warning(f"매칭 실패 항목 {len(unmatched)}개 (마스터 파일에 없는 이름/규격일 수 있습니다):")
                st.text("\n".join(unmatched[:10])) # 10개까지만 표시

            output = BytesIO()
            wb.save(output)
            output.seek(0)
            st.download_button("수정된 파일 다운로드", data=output, file_name=f"매칭완료_{uploaded_file.name}")
    except Exception as e:
        st.error(f"오류: {e}")
