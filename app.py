# ... (앞부분 생략)

            else:
                count = 0
                # 수직 중앙 정렬(vertical='center') 및 수평 우측 정렬(horizontal='right') 추가
                # number_fmt는 그대로 유지
                cell_style = Alignment(horizontal='right', vertical='center') 
                number_fmt = '#,##0'

                for row in range(start_row, ws.max_row + 1):
                    val_a = str(ws.cell(row=row, column=c_name).value or "").strip()
                    val_b = str(ws.cell(row=row, column=c_spec).value or "").strip()
                    key = f"{val_a}_{val_b}"
                    
                    if key in master_data:
                        # 1. 재료비 처리
                        cell_mat = ws.cell(row=row, column=c_mat)
                        cell_mat.value = float(str(master_data[key]['E']).replace(',', ''))
                        cell_mat.number_format = number_fmt
                        cell_mat.alignment = cell_style
                        
                        # 2. 노무비 처리
                        cell_lab = ws.cell(row=row, column=c_lab)
                        cell_lab.value = float(str(master_data[key]['G']).replace(',', ''))
                        cell_lab.number_format = number_fmt
                        cell_lab.alignment = cell_style
                        
                        # 3. 경비 처리
                        cell_exp = ws.cell(row=row, column=c_exp)
                        cell_exp.value = float(str(master_data[key]['I']).replace(',', ''))
                        cell_exp.number_format = number_fmt
                        cell_exp.alignment = cell_style
                        
                        count += 1
                
                st.success(f"완료! 총 {count}개의 항목에 단가가 입력되었습니다.")
                
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                final_name = f"매칭완료_{uploaded_file.name.split('.')[0]}.{file_ext}"
                st.download_button(label=f"수정된 파일 다운로드 ({file_ext})", data=output, file_name=final_name)
            
# ... (뒷부분 생략)
