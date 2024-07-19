import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import datetime

# Tkinter 창 열기
Tk().withdraw()

# 사용자에게 엑셀 파일 선택하도록 요청
file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

# 엑셀 파일에서 모든 시트 읽어오기
xls = pd.ExcelFile(file_path)

# 날짜 및 플랫폼 값 입력 받기
target_date = input("날짜(MMDD ex - 0409)를 입력하세요: ")
platform = input("플랫폼(必) 값을 입력하세요: ")

# 날짜 형식이 'MMDD'일 경우 'YYYY-MM-DD' 형식으로 변환
if len(target_date) == 4:
    target_date = '2024-' + target_date[:2] + '-' + target_date[2:]

# 모든 시트의 데이터프레임을 저장할 리스트 생성
all_dfs = []

# 각 시트의 이름을 가져와서 처리
for sheet_name in xls.sheet_names:
    # 각 시트에서 데이터 읽기
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # 구매일 열 찾기
    purchase_date_col = None
    for col in df.columns:
        if '구매일' in str(df[col].iloc[2]):  # 3번째 행에 '구매일'이 들어있는 열을 찾습니다.
            purchase_date_col = col
            break

    if purchase_date_col is not None:
        # '구매일' 열을 datetime 형식으로 변환 (두 번째 행부터 시작하여 변환)
        df['구매일'] = pd.to_datetime(df[purchase_date_col].iloc[3:])
        
        # 해당 날짜에 해당하는 행만 필터링
        df_selected_date = df[df['구매일'].dt.strftime('%Y-%m-%d') == target_date]
        
        # 새로운 헤더 값을 설정
        new_header = df.iloc[2].values
        df_selected_date.columns = new_header
        
        df_selected_date = df_selected_date.reset_index(drop=True)
        
        # 새로운 형식의 시트를 담을 빈 데이터프레임 생성
        new_df = pd.DataFrame(columns=['순번', '수취인(必)', '연락처(必)', '연락처2(선택기입사항)', '주소(必)', '우편번호(작성x)', 
                                        '배송메모(작성x)', '플랫폼(必)', '출고상품명or스토어명[必]'])

        # 필요한 정보 추출하여 새로운 형식의 데이터프레임에 추가
        new_df['순번'] = range(1, len(df_selected_date) + 1)
        new_df['수취인(必)'] = df_selected_date['수취인']
        new_df['연락처(必)'] = df_selected_date['연락처']
        new_df['연락처2(선택기입사항)'] = ''
        new_df['주소(必)'] = df_selected_date['배송지']
        new_df['우편번호(작성x)'] = ''
        new_df['배송메모(작성x)'] = ''
        new_df['플랫폼(必)'] = platform
        new_df['출고상품명or스토어명[必]'] = sheet_name
        
        # 생성된 데이터프레임을 리스트에 추가
        all_dfs.append(new_df)
        
    else:
        print(f"'{sheet_name}' 시트에서 구매일 열을 찾을 수 없습니다.")

# 리스트에 저장된 데이터프레임을 하나로 합치기
final_df = pd.concat(all_dfs, ignore_index=True)
final_df['순번'] = range(1, len(final_df) + 1)

# 파일 이름 설정
file_name = f"{target_date}_{len(final_df)}"

# 데이터프레임을 엑셀 파일로 저장
with pd.ExcelWriter(f"{file_name}건요청.xlsx", engine='xlsxwriter') as writer:
    final_df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=1, header=False)
    
    # 엑셀 시트 포맷 설정
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # 열 너비 자동 조정
    for i, col in enumerate(final_df.columns):
        max_width = max(final_df[col].astype(str).map(len).max(), len(col))
        header_width = max(max_width, len(col))
        worksheet.set_column(i, i, max_width)
        
    # 헤더 추가
    for col_num, value in enumerate(final_df.columns.values):
        worksheet.write(0, col_num, value)

    border_format = workbook.add_format({'border': 1, 'border_color': 'black'})
    for row_num, row_data in final_df.iterrows():
        for col_num, value in enumerate(row_data):
            worksheet.write(row_num + 1, col_num, value, border_format)
   # 헤더에 굵은 테두리 추가
    header_border_format = workbook.add_format({'border': 2, 'border_color': 'black', 'bold': True})
    for col_num, value in enumerate(final_df.columns.values):
        worksheet.write(0, col_num, value, header_border_format)         
