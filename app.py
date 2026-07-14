import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import traceback

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
        # header=None 추가: 첫 줄을 제목으로 보지 않고 데이터로 읽어들임
        df = pd.read_csv(GOOGLE_SHEET_URL, header=None)
        
        if df.empty:
            return "시트에 데이터가 없습니다."
        
        master_data = {}
        for _, row in df.iterrows():
            # 첫번째(A), 두번째(B) 열로 키 생성 (0번, 1번 인덱스)
            name = str(row[0]).strip() if pd.notnull(row[0]) else ""
            spec = str(row[1]).strip() if pd.notnull(row[1]) else ""
            key = f"{name}_{spec}"
            
            # E(5번째열=인덱스4), G(7번째열=인덱스6), I(9번째열=인덱스8)
            master_data[key] = {
                'E': row[4] if len(row) > 4 else None,
                'G': row[6] if len(row) > 6 else None,
                'I': row[8] if len(row) > 8 else None
            }
        return master_data
    except Exception as e:
        return f"Error: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"

# 데이터 로드 실행
result = load_master_data()

if isinstance(result, str):
    st.error(f"마스터 데이터를 불러오는 중 오류 발생: {result}")
    master_data = None
else:
    master_data = result

# 4. 파일 업로드 및 시트 선택
uploaded_file = st.file_uploader("작업할 내역서 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file and master_data:
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        sheet_names = wb.sheetnames
        
        selected_sheet = st.selectbox("수정할 시트를 선택하세요", sheet_names)
        ws = wb[selected_sheet]

        if st.button("단가 자동 입력 실행"):
            count = 0
            for row in range(2, ws.max_row + 1): 
                val_a = ws.cell(row=row, column=1).value
                val_b = ws.cell(row=row, column=2).value
                
                name = str(val_a).strip() if val_a is not None else ""
                spec = str(val_b).strip() if val_b is not None else ""
                current_key = f"{name}_{spec}"
                
                if current_key in master_data:
                    # E(5), G(7), I(9) 열 값 수정
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
    except Exception as e:
        st.error(f"엑셀 처리 중 오류 발생: {str(e)}")
