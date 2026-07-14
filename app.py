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

# 3. 마스터 파일 로드
try:
    # 깃허브에 올려둔 마스터 파일 (master_price.xlsx)을 불러옵니다.
    master_wb = openpyxl.load_workbook('master_price.xlsx', data_only=True)
    master_ws = master_wb.active
    
    # 마스터 데이터를 딕셔너리로 변환 (빠른 검색을 위해)
    # 예: A열(품명)을 키(key)로, E, G, I열 값을 밸류(value)로 저장한다고 가정합니다.
    # [주의] 이 부분은 실제 엑셀 양식에 맞춰 수정이 필요합니다.
    master_data = {}
    for row in range(2, master_ws.max_row + 1):
        key = str(master_ws.cell(row=row, column=1).value) # 1열(A)을 기준으로 매칭
        master_data[key] = {
            'E': master_ws.cell(row=row, column=5).value, # 5열(E)
            'G': master_ws.cell(row=row, column=7).value, # 7열(G)
            'I': master_ws.cell(row=row, column=9).value  # 9열(I)
        }
except Exception as e:
    st.error(f"마스터 파일을 불러오는 중 오류가 발생했습니다: {e}")
    st.stop()

# 4. 파일 업로드 및 시트 선택
uploaded_file = st.file_uploader("작업할 내역서 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    # 엑셀 파일 열기 (수식 유지)
    wb = openpyxl.load_workbook(uploaded_file)
    sheet_names = wb.sheetnames
    
    selected_sheet = st.selectbox("수정할 시트를 선택하세요", sheet_names)
    ws = wb[selected_sheet]

    if st.button("단가 자동 입력 실행"):
        # E(5), G(7), I(9) 열 수정 로직
        count = 0
        for row in range(2, ws.max_row + 1): # 2행부터 데이터가 있다고 가정
            item_name = str(ws.cell(row=row, column=1).value) # A열 기준으로 검색
            
            if item_name in master_data:
                # E, G, I 열만 수정 (다른 열과 수식은 건드리지 않음)
                ws.cell(row=row, column=5).value = master_data[item_name]['E'] # E열
                ws.cell(row=row, column=7).value = master_data[item_name]['G'] # G열
                ws.cell(row=row, column=9).value = master_data[item_name]['I'] # I열
                count += 1
        
        st.success(f"성공! {count}개의 항목이 업데이트되었습니다.")

        # 파일 저장
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        # 다운로드 버튼
        st.download_button(
            label="수정된 파일 다운로드",
            data=output,
            file_name=f"수정본_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
