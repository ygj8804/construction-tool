# --- [탭 1] 단가 자동 입력 (서식 유지 + 데이터 매칭) ---
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
        wb = openpyxl.load_workbook(uploaded_file)
        ws_name = st.selectbox("작업할 시트 선택", wb.sheetnames, key="sheet1")
        ws = wb[ws_name]
        
        if st.button("단가 매칭 실행", type="primary"):
            # 1. 마스터 데이터 불러오기
            master_df = pd.read_csv(DATA_URL)
            
            # 2. 엑셀 파일 데이터 읽기 (Pandas로 변환)
            data = ws.values
            df_ex = pd.DataFrame(data[start_row-1:], columns=data[start_row-1])
            
            # 3. 매칭 로직 (품명+규격 기준)
            for i, row in df_ex.iterrows():
                row_idx = start_row + i
                name, spec = row[col_name], row[col_spec]
                
                match = master_df[(master_df['품명'] == name) & (master_df['규격'] == spec)]
                
                if not match.empty:
                    ws[f"{col_mat}{row_idx}"] = match['재료비'].values[0]
                    ws[f"{col_lab}{row_idx}"] = match['노무비'].values[0]
                    ws[f"{col_exp}{row_idx}"] = match['경비'].values[0]
            
            # 4. 결과 저장 및 다운로드 (서식 유지)
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success("🎉 단가 매칭이 완료되었습니다! 아래 버튼을 눌러 결과 파일을 받으세요.")
            st.download_button(
                label="결과 파일 다운로드 (서식 유지)",
                data=output,
                file_name=f"매칭완료_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            wb.close()
