import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# 1. 페이지 설정
st.set_config = st.set_page_config(page_title="서원건설 단가 자동 입력기", layout="wide")

# 2. 헤더 및 작성자 표시
st.title("🏗️ 서원건설")
st.subheader("공사 내역서 단가 자동 입력기")
st.markdown("**만든이: 유강진 대리**")
st.markdown("---")

# 3. 설정 창 (상단 배치 - 유 대리님 요청 양식 반영)
with st.expander("⚙️ 내역서 열 설정 (설정을 변경하면 매칭 기준이 바뀝니다)", expanded=True):
    st.write("사용하시는 엑셀 양식에 맞춰 열 번호를 입력하세요.")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        start_row = st.number_input("데이터 시작 행", value=2, min_value=1)
        col_name = st.number_input("품명(A) 열 번호", value=1, min_value=1)
        col_spec = st.number_input("규격(B) 열 번호", value=2, min_value=1)
        
    with col2:
        col_unit = st.number_input("단위(C) 열 번호", value=3, min_value=1)
        col_mat = st.number_input("재료비단가(E) 열 번호", value=5, min_value=1)
        
    with col3:
        col_lab = st.number_input("노무비단가(G) 열 번호", value=7, min_value=1)
        col_exp = st.number_input("경비단가(I) 열 번호", value=9, min_value=1)

# 4. 구글 시트(마스터 데이터) 로드
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"

@st.cache_data(ttl=60) 
def load_master_data():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL, header=None)
        if df.empty: return None
        
        master_data = {}
        for _, row in df.iterrows():
            # A(0) + B(1)을 키로 사용 (품명+규격)
            name = str(row[0]).strip() if pd.notnull(row[0]) else ""
            spec = str(row[1]).strip() if pd.notnull(row[1]) else ""
            key = f"{name}_{spec}"
            
            # E(4), G(6), I(8) 인덱스 매칭
            master_data[key] = {
                'E': row[4] if len(row) > 4 else None,
                'G': row[6] if len(row) > 6 else None,
                'I': row[8] if len(row) > 8 else None
            }
        return master_data
    except Exception as e:
        return str(e)

master_data = load_master_data()

# 5. 파일 업로드 및 실행
uploaded_file = st.file_uploader("작업할 내역서 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file and master_data:
    if isinstance(master_data, str):
        st.error(f"데이터 로드 오류: {master_data}")
    else:
        try:
            wb = openpyxl.load_workbook(uploaded_file)
            selected_sheet = st.selectbox("수정할 시트를 선택하세요", wb.sheetnames)
            ws = wb[selected_sheet]

            if st.button("단가 자동 입력 실행"):
                count = 0
                # 설정된 시작 행부터 끝까지 반복
                for row in range(start_row, ws.max_row + 1): 
                    # 사용자 입력값(열 번호)을 그대로 사용하여 읽음
                    val_name = ws.cell(row=row, column=col_name).value
                    val_spec = ws.cell(row=row, column=col_spec).value
                    
                    name = str(val_name).strip() if val_name is not None else ""
                    spec = str(val_spec).strip() if val_spec is not None else ""
                    current_key = f"{name}_{spec}"
                    
                    if current_key in master_data:
                        # 설정된 열(E, G, I)에 값 대입
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
                
                st.download_button(
                    label="수정된 파일 다운로드",
                    data=output,
                    file_name=f"수정본_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"엑셀 처리 중 오류 발생: {str(e)}")
