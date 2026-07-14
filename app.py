import streamlit as st
import openpyxl
import io
import os
import pandas as pd
import unicodedata

MASTER_FILE_NAME = "최저단가_기준_마스터.xlsx"

# --- [정규화 함수] ---
def normalize_string(s):
    if s is None: return ""
    s = unicodedata.normalize('NFKC', str(s))
    s = s.strip().replace(" ", "").replace("\n", "").replace("\r", "").replace("\u00A0", "").lower()
    return s

def col_letter_to_index(letter):
    try: return openpyxl.utils.column_index_from_string(letter.upper()) - 1
    except: return None

# --- [UI 디자인] ---
st.set_page_config(page_title="서원건설 단가 입력기", layout="centered", page_icon="🏗️")

with st.sidebar:
    st.header("⚙️ 설정 및 관리")
    uploaded_master = st.file_uploader("최저단가 마스터 엑셀 업로드", type=["xlsx"])
    if uploaded_master:
        with open(MASTER_FILE_NAME, "wb") as f: f.write(uploaded_master.getbuffer())
        st.success("마스터 파일 적용됨.")

st.title("🚀 공사 내역서 단가 자동 입력기")

# 마스터 로드
master_dict = {}
if os.path.exists(MASTER_FILE_NAME):
    try:
        wb = openpyxl.load_workbook(MASTER_FILE_NAME, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=4, values_only=True):
            if not row[0]: continue
            key = (normalize_string(row[0]), normalize_string(row[1]), normalize_string(row[2]))
            master_dict[key] = {'mat': row[4], 'lab': row[6], 'exp': row[8]}
        st.success(f"✅ 마스터 데이터 로드 완료 ({len(master_dict)}개 품목)")
    except Exception as e:
        st.error(f"마스터 파일 읽기 오류: {e}")

if master_dict:
    with st.expander("📌 설정 (열 위치 및 시작 행)"):
        col1, col2, col3 = st.columns(3)
        in_name = col1.text_input("품목명 열", "B")
        in_spec = col2.text_input("규격 열", "C")
        in_unit = col3.text_input("단위 열", "D")
        
        col4, col5, col6 = st.columns(3)
        t_mat = col4.text_input("재료비 열", "E")
        t_lab = col5.text_input("노무비 열", "G")
        t_exp = col6.text_input("경비 열", "I")
        start_row = st.number_input("데이터 시작 행", min_value=1, value=25)

    debug_mode = st.checkbox("🔍 [디버그 모드] 매칭 실패 품목 보기")
    target_file = st.file_uploader("단가 채울 견적서 엑셀 업로드", type=["xlsx", "xls"])
    
    if target_file:
        try:
            # 1. 파일 읽기 및 시트 선택
            if target_file.name.endswith('.xls'):
                df = pd.read_excel(target_file, engine='xlrd', sheet_name=None)
                sheet_names = list(df.keys())
                selected_sheet = st.selectbox("작업할 시트를 선택하세요", sheet_names)
                df_selected = pd.read_excel(target_file, engine='xlrd', sheet_name=selected_sheet)
                buffer = io.BytesIO()
                df_selected.to_excel(buffer, index=False)
                buffer.seek(0)
                wb = openpyxl.load_workbook(buffer)
                ws = wb.active
            else:
                wb = openpyxl.load_workbook(target_file, data_only=True)
                sheet_names = wb.sheetnames
                selected_sheet = st.selectbox("작업할 시트를 선택하세요", sheet_names)
                ws = wb[selected_sheet]

            # [진단 기능] 파일 정보 출력
            st.info(f"📂 파일명: {target_file.name} | 선택 시트: {selected_sheet}")
            st.write(f"📊 **파일 내 총 행 수:** {ws.max_row}행")
            
            n_idx = col_letter_to_index(in_name)
            s_idx = col_letter_to_index(in_spec)
            u_idx = col_letter_to_index(in_unit)
            m_idx = [col_letter_to_index(c) for c in [t_mat, t_lab, t_exp]]
            
            if st.button("단가 매칭 시작"):
                match = 0
                with st.spinner("단가 매칭 중..."):
                    for i, row in enumerate(ws.iter_rows(min_row=start_row), start_row):
                        if n_idx >= len(row) or not row[n_idx].value: continue
                        
                        val_name = row[n_idx].value
                        val_spec = row[s_idx].value if s_idx < len(row) else ""
                        val_unit = row[u_idx].value if u_idx < len(row) else ""
                        
                        key = (normalize_string(val_name), normalize_string(val_spec), normalize_string(val_unit))
                        
                        if key in master_dict:
                            row[m_idx[0]].value = master_dict[key]['mat']
                            row[m_idx[1]].value = master_dict[key]['lab']
                            row[m_idx[2]].value = master_dict[key]['exp']
                            match += 1
                        elif debug_mode:
                            st.write(f"❌ 매칭 실패 (행 {i}): {val_name} / {val_spec} / {val_unit}")
                
                st.success(f"🎉 작업 완료! 총 {match}개 품목 단가 입력됨.")
                base_name = os.path.splitext(target_file.name)[0]
                output = io.BytesIO()
                wb.save(output)
                st.download_button("📥 결과 파일 다운로드", output.getvalue(), f"단가완료_{base_name}.xlsx")
        except Exception as e:
            st.error(f"오류: {e}")