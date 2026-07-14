import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")

# 2. 헤더 및 작성자 표시
st.title("🏗️ 서원건설")
st.subheader("공사 내역서 단가 자동 입력기")
st.sidebar.markdown("---")
st.sidebar.markdown("**만든이: 유강진 대리**")

# 3. 구글 시트(마스터 데이터) 실시간 로드
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

@st.cache_data(ttl=60) 
def load_master_data():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL)
        # 품명(A열) + 규격(B열)을 합쳐서 하나의 고유 키로 생성
        master_data = {}
        for _, row in df.iterrows():
            # A열(인덱스 0)과 B열(인덱스 1)을 합쳐서 키 생성
            name = str(row[0]).strip() if pd.notnull(row[0]) else ""
            spec = str(row[1]).strip() if pd.notnull(row[1]) else ""
            key = f"{name}_{spec}"
            
            master_data[key] = {
                'E': row[4] if pd.notnull(row[4]) else None, # 5열(E)
                'G': row[6] if pd.notnull(row[6]) else None, # 7열(G)
                'I': row[8] if pd.notnull(row[8]) else None  # 9열(I)
            }
        return master_data
    except Exception as e:
        st.error(f"마스터 데이터를 불러오는 중 오류 발생: {e}")
        return None

master_data = load_master_data()

# 4. 파일 업로드 및 시트 선택
uploaded_file = st.file_uploader("작업할 내역서 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file and master_data:
    wb = openpyxl.load_workbook(uploaded_file)
    sheet_names = wb.sheetnames
    
    selected_sheet = st.selectbox("수정할 시트를 선택하세요", sheet_names)
    ws = wb[selected_sheet]

    if st.button("단가 자동 입력 실행"):
        count = 0
        # A열과 B열을 읽어서 마스터 데이터 키와 비교
        for row in range(2, ws.max_row + 1): 
            val_a = ws.cell(row=row, column=1).value # A열
            val_b = ws.cell(row=row, column=2).value # B열
            
            name = str(val_a).strip() if val_a is not None else ""
            spec = str(val_b).strip() if val_b is not None else ""
            current_key = f"{name}_{spec}"
            
            if current_key in master_data:
                # E(5), G(7), I(9) 열만 값 수정
                if master_data[current_key]['E'] is not None:
                    ws.cell(row=row, column=5).value = master_data[current_key]['E']
                if master_data[current_key]['G'] is not None:
                    ws.cell(row=row, column=7).value = master_data[current_key]['G']
                if master_data[current_key]['I'] is not None:
                    ws.cell(row=row, column=9).value = master_data[current_key]['I']
                count += 1
        
        st.success(f"성공! {count}개의 항목이 업데이트되었습니다.")

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="수정된 파일 다운로드",
            data=output,
            file_name=f"수정본_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
