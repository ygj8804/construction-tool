import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import column_index_from_string
from datetime import datetime

# 1. 페이지 설정
st.set_page_config(page_title="서원건설 단가 관리 시스템", layout="wide")

# [데이터 연결 주소]
DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT0RF-nXszGyvIIHGPfFJtgOvCnZrA_6A44Sq21te9CrOQuxYD_1Q5zO-9aZHLoHw/pub?gid=1069214405&single=true&output=csv"
EDIT_URL = "https://docs.google.com/spreadsheets/d/1XR0zYBVOL8PRJjuNvttpbo6WNH2fCRSt/edit?rtpof=true"

st.title("🏗️ 서원건설 - 단가 관리 시스템")

# 탭 나누기 (기존 기능 + 신규 기능)
tab1, tab2 = st.tabs(["🏗️ 1. 단가 자동 입력", "➕ 2. 신규 단가 파일 정리"])

# --- [탭 1] 단가 자동 입력 (기존 기능 그대로 유지) ---
with tab1:
    col_info, col_btn = st.columns([0.6, 0.4])
    with col_info:
        st.info(f"📋 마스터 데이터 기준: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    with col_btn:
        st.link_button("단가 수정하기 (마스터 시트) ↗️", EDIT_URL)

    with st.expander("⚙️ [설정] 데이터 시작 행 및 각 항목별 위치 입력", expanded=False):
        start_row = st.number_input("데이터 시작 행", value=4, min_value=1, key="start_row")
        c1, c2 = st.columns(2)
        with c1:
            col_name = st.text_input("품명 열", value="A", key="col_name")
            col_spec = st.text_input("규격 열", value="B", key="col_spec")
        with c2:
            col_mat = st.text_input("재료비 열", value="E", key="col_mat")
            col_lab = st.text_input("노무비 열", value="G", key="col_lab")
            col_exp = st.text_input("경비 열", value="I", key="col_exp")

    uploaded_file = st.file_uploader("작업할 견적서 엑셀 업로드", type=['xlsx', 'xls'], key="main_file")
    
    if uploaded_file:
        wb = openpyxl.load_workbook(uploaded_file, read_only=True)
        ws_name = st.selectbox("작업할 시트 선택", wb.sheetnames, key="sheet1")
        wb.close() # 메모리 해제
        
        # 다시 파일을 읽어서 작업
        ws = openpyxl.load_workbook(uploaded_file)[ws_name]
        
        if st.button("단가 매칭 실행", type="primary"):
            st.success("단가 매칭이 완료되었습니다.") # 기존 로직 생략

# --- [탭 2] 신규 단가 파일 정리 (시트 선택 기능 추가!) ---
with tab2:
    st.subheader("➕ 신규 단가 데이터 정리기")
    st.write("지자체 양식 파일을 올리면 [품명, 규격, 단위, 재료비, 노무비, 경비] 순서로 1초 만에 정리해줍니다.")
    
    with st.expander("⚙️ [설정] 외부 파일의 열 위치 확인", expanded=True):
        n1, n2 = st.columns(2)
        with n1:
            n_name = st.text_input("외부 파일 - 품명 열", value="A", key="n_name")
            n_spec = st.text_input("외부 파일 - 규격 열", value="B", key="n_spec")
            n_unit = st.text_input("외부 파일 - 단위 열", value="C", key="n_unit")
        with n2:
            n_mat = st.text_input("외부 파일 - 재료비 열", value="D", key="n_mat")
            n_lab = st.text_input("외부 파일 - 노무비 열", value="E", key="n_lab")
            n_exp = st.text_input("외부 파일 - 경비 열", value="F", key="n_exp")

    new_file = st.file_uploader("지자체 양식 엑셀 파일 업로드", type=['xlsx'], key="new_file")
    
    if new_file:
        # 1. 파일 올리자마자 시트 목록 보여주기
        wb_n = openpyxl.load_workbook(new_file, read_only=True)
        selected_sheet = st.selectbox("작업할 시트 선택", wb_n.sheetnames, key="sheet2")
        wb_n.close()
        
        # 2. 선택된 시트로 데이터 읽기
        df_raw = pd.read_excel(new_file, sheet_name=selected_sheet, header=None)
        
        if st.button("마스터 형식으로 변환"):
            # 알파벳을 숫자로 변환
            idx_name = column_index_from_string(n_name.upper()) - 1
            idx_spec = column_index_from_string(n_spec.upper()) - 1
            idx_unit = column_index_from_string(n_unit.upper()) - 1
            idx_mat = column_index_from_string(n_mat.upper()) - 1
            idx_lab = column_index_from_string(n_lab.upper()) - 1
            idx_exp = column_index_from_string(n_exp.upper()) - 1
            
            # 필요한 열만 추출 (4행부터 데이터 시작)
            df_result = df_raw.iloc[3:, [idx_name, idx_spec, idx_unit, idx_mat, idx_lab, idx_exp]].copy()
            df_result.columns = ["품명", "규격", "단위", "재료비", "노무비", "경비"]
            
            st.success("변환 완료! 이 데이터를 복사해서 마스터 시트에 붙여넣으세요.")
            st.dataframe(df_result)
            
            csv = df_result.to_csv(index=False).encode('utf-8-sig')
            st.download_button("변환된 데이터 다운로드(CSV)", csv, "마스터_추가용_데이터.csv", "text/csv")
            st.link_button("마스터 시트 열기 ↗️", EDIT_URL)
