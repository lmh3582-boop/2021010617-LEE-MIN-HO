import pandas as pd
import os

# --- 라이브러리 임포트 ---
# pandas: 데이터 처리 (CSV 읽기, 정렬 등)
# os: 파일 경로 및 이름 처리

try:
    # --- (1) 파일 입력 및 설정 ---
    
    # 사용자에게 CSV 파일명 입력받기 (예: 0.9)
    input_file = input("Enter the CSV file name: ").strip()
    
    # 입력값 끝에 .csv가 없으면 자동으로 추가
    if not input_file.lower().endswith('.csv'):
        input_file += '.csv'
    
    # 파일명에서 확장자(.csv)를 제거 (예: '0.9')
    base_name = os.path.splitext(input_file)[0]
    # 최종 엑셀 파일 이름 설정 (예: 0.9.xlsx)
    output_excel_file = base_name + ".xlsx"
    
    # 판다스로 CSV 파일 읽기
    df = pd.read_csv(input_file)
    print(f"Successfully loaded '{input_file}' file.")
    
    # (수정 가능) CSV에서 읽어올 필수 열 이름
    required_columns = ['Points_0', 'Points_2', 'Pressure_Coefficient']
    
    # --- (2) 데이터 처리 및 정규화 ---
    
    # 필수 열이 파일에 모두 있는지 확인
    if all(col in df.columns for col in required_columns):
        # 원본을 냅두고, 필수 열만 복사해서 사용
        df_processed = df[required_columns].copy()
        
        # Points_0 (X좌표)의 최소값/최대값 찾기
        min_x = df_processed['Points_0'].min()
        max_x = df_processed['Points_0'].max()
        
        # (안정성) X좌표가 모두 같으면 0으로 나누기 방지
        if (max_x - min_x) == 0:
            df_processed['Points_0_Normalized'] = 0.5
        else:
            # X좌표 정규화 (0 ~ 1 사이 값으로 변경)
            df_processed['Points_0_Normalized'] = (df_processed['Points_0'] - min_x) / (max_x - min_x)
            
        # Points_2 (Y좌표) 기준으로 상/하부 표면 분리 (0은 상부에 포함)
        df_upper = df_processed[df_processed['Points_2'] >= 0].copy()
        df_lower = df_processed[df_processed['Points_2'] < 0].copy()
        
        # 상부 표면: X좌표(Normalized) 기준 오름차순 정렬 (0 -> 1)
        upper_sorted = df_upper[['Points_0_Normalized', 'Pressure_Coefficient']].sort_values(
            by='Points_0_Normalized', ascending=True
        )
        
        # 하부 표면: X좌표(Normalized) 기준 내림차순 정렬 (1 -> 0)
        lower_sorted = df_lower[['Points_0_Normalized', 'Pressure_Coefficient']].sort_values(
            by='Points_0_Normalized', ascending=False
        )

        # 그래프를 닫기 위해, 상부 표면의 첫 번째 점(정체점)을 따로 저장
        if not upper_sorted.empty:
            closing_point = upper_sorted.iloc[0:1].copy()
        else:
            # (안정성) 상부 표면이 비어있으면 빈 DataFrame 생성
            closing_point = pd.DataFrame(columns=upper_sorted.columns) 

        # 상부(0->1), 하부(1->0), 닫는점(0) 순서로 데이터 합치기
        data_frames_to_concat = [upper_sorted, lower_sorted]
        if not closing_point.empty:
            data_frames_to_concat.append(closing_point)
        
        combined_data_closed = pd.concat(data_frames_to_concat, ignore_index=True)

        # --- (3) 엑셀 파일 생성 (xlsxwriter 엔진 사용) ---
        
        # pandas ExcelWriter 객체 생성
        writer = pd.ExcelWriter(output_excel_file, engine='xlsxwriter')
        
        # (수정 가능) 엑셀 시트 이름
        sheet_name = 'Cp_Data'
        # 정렬된 데이터를 'Cp_Data' 시트에 쓰기 (인덱스 제외)
        combined_data_closed.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 엑셀 파일과 시트 객체 가져오기 (차트 삽입용)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # 엑셀 차트 객체 생성 (산점도, 부드러운 선)
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        
        # 데이터 행 개수 세기 (차트 범위 설정용)
        num_rows = len(combined_data_closed)
        
        # 엑셀의 어느 범위(셀)를 차트 데이터로 쓸지 지정
        # 예: '=Cp_Data!$A$2:$A$93'
        x_axis_data = f'={sheet_name}!$A$2:$A${num_rows + 1}'
        y_axis_data = f'={sheet_name}!$B$2:$B${num_rows + 1}'
        
        # 차트에 실제 데이터(계열) 추가
        chart.add_series({
            'categories': x_axis_data,  # X축 데이터
            'values':     y_axis_data,  # Y축 데이터
            'line':       {'color': 'black', 'width': 2}, # (수정 가능) 선 서식
            'marker':     {'type': 'none'}  # 데이터 점 표식(마커) 없음
        })
        
        # --- (4) 차트 서식 설정 ---

        # 1. 동적 차트 제목 생성 (파일명 기준)
        try:
            # 파일명(base_name)을 숫자로 바꿔서 제목 생성 (예: 0.9 -> Section 90%)
            percent_value = int(float(base_name) * 100)
            chart_title_text = f"Section {percent_value}%"
        except ValueError:
            # (안정성) 파일명이 숫자가 아니면 (예: 'test') 기본 제목 사용
            chart_title_text = "Pressure Coefficient vs. x/c"

        # 2. (수정 가능) 글꼴 서식 사전 정의
        title_font = {'name': '맑은 고딕', 'size': 14, 'bold': True}
        axis_font = {'name': '맑은 고딕', 'size': 14, 'bold': True}

        # 3. 차트 제목 적용
        chart.set_title({
            'name': chart_title_text,
            'name_font': title_font
        })

        # 4. X축 설정
        chart.set_x_axis({
            'name': 'X/C',  # (수정) X축 제목을 다시 원래대로 설정
            'name_font': axis_font,
            'major_gridlines': {  # 세로 눈금선
                'visible': True,
                'line': {'color': '#D9D9D9', 'dash_type': 'dash'} # (수정 가능) 눈금선 서식
            },
            'label_position': 'high' # X축 레이블을 하단에 고정 (Y축 reverse시 'high'가 하단)
        })
        
        # 5. Y축 설정
        chart.set_y_axis({
            'name': 'Cp',  # (수정 가능) Y축 제목
            'name_font': axis_font,
            'reverse': True,  # Y축 뒤집기 (Cp 그래프 관례)
            'major_gridlines': {  # 가로 눈금선
                'visible': True,
                'line': {'color': '#D9D9D9'} # (수정 가능) 눈금선 서식
            }
        })
        
        # 6. (수정) 범례(Legend)를 다시 숨깁니다.
        chart.set_legend({'position': 'none'})
        
        # (수정 가능) 시트의 'D2' 셀 위치에 차트 삽입 및 크기 조절
        worksheet.insert_chart('D2', chart, {'x_scale': 1.5, 'y_scale': 1.5})
        
        # --- (5) 엑셀 파일 저장 및 닫기 ---
        writer.close()
        
        print(f"\nSuccess! Data and chart saved to '{output_excel_file}'")

    # --- (6) 예외 처리 ---
    else:
        # (예외 처리) CSV에 필수 열이 없을 경우
        print(f"Error: Required columns {required_columns} not found in {input_file}.")

except FileNotFoundError:
    # (예외 처리) 파일이 없을 경우
    print(f"Error: '{input_file}' file not found. Make sure the file is in the same directory as the script.")
except ImportError:
    # (예외 처리) xlsxwriter 라이브러리가 설치되지 않았을 경우
    print("\nError: 'xlsxwriter' library not found.")
    print("Please install it by running: pip install xlsxwriter")
except Exception as e:
    # (예외 처리) 그 외 모든 알 수 없는 에러
    print(f"An error occurred during data processing: {e}")