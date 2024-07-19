import subprocess
import sys
from importlib import util
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import pandas as pd
import re
from collections import Counter
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Google API 사용을 위한 인증 정보 설정
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)  # Correct way to create credentials
client = gspread.authorize(creds)

# 구글 스프레드시트 및 시트 열기
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1TZU6t5sT2wYfN1KwdIo2jBiSfp1fq7MA6Byxbd3KrIM/edit#gid=0")
sheet = spreadsheet.worksheet('시트1')

# DataFrame을 Google Sheets에 업로드
def upload_to_google_sheets(df, sheet):
    # NaN 값을 빈 문자열로 대체
    df = df.replace([float('inf'), float('-inf'), None], '')  # inf 값 처리 추가
    df = df.fillna('')

    # 기존 데이터를 지웁니다
    sheet.clear()
    
    # 데이터프레임을 리스트로 변환하여 업데이트합니다
    rows = [df.columns.values.tolist()] + df.values.tolist()
    
    # 시작 셀을 B12로 지정
    start_cell = 'B19'
    
    # 시작 셀에 데이터 업데이트
    sheet.update(start_cell, rows)
    
# 웹 드라이버 옵션 설정
options = Options()
options.add_argument("--headless")  # 웹 브라우저 화면을 표시하지 않음

# 웹 드라이버 실행
driver = webdriver.Chrome(options=options)

# 웹사이트로 이동
url = "https://xn--ok-vd0j004f7nc.com/bbs/login.php?url=https%3A%2F%2Fxn--ok-vd0j004f7nc.com%2Fadm"
driver.get(url)

# 로그인 정보 입력
mb_id = driver.find_element(By.NAME, "mb_id")
mb_password = driver.find_element(By.NAME, "mb_password")

mb_id.send_keys("admin")
mb_password.send_keys("qwercmcm!@#$")

# 로그인 버튼 클릭
login_button = driver.find_element(By.CSS_SELECTOR,
                                   "#thema_wrapper > div.at-body > div > div.row > div > div > div.form-body > form > div.row > div:nth-child(2) > button")
login_button.click()

# 회원관리 메뉴 클릭
member_management_link = driver.find_element(By.LINK_TEXT, "회원관리")
member_management_link.click()

# 접속자검색 메뉴 클릭
visit_search_link = driver.find_element(By.LINK_TEXT, "접속자검색")
visit_search_link.click()

# 시작일과 끝일을 입력받아서 그 사이의 모든 날짜 생성
start_date_input = input("검색할 시작 날짜를 입력하세요(예: 1023): ")
end_date_input = input("검색할 종료 날짜를 입력하세요(예: 1025): ")

# 시작일과 끝일을 datetime 객체로 변환
start_date = datetime.strptime(start_date_input, "%m%d")
end_date = datetime.strptime(end_date_input, "%m%d")

# 시작일부터 끝일까지의 모든 날짜를 생성하여 리스트에 저장
date_numbers = [(start_date + timedelta(days=i)).strftime("%m%d") for i in range((end_date - start_date).days + 1)]
print("검색할 날짜:", date_numbers)


# 모든 데이터를 저장할 리스트 초기화
all_data = []
all_dates = []

data_dict = {
    '223298350263': '인천간판 OK디자인 간판 종류와 시공 과정',
    '223250711888': '인천간판 실외간판 알아본다면',
    '223112503558': '인천간판 ok디자인 종류 알아보기',
    '223291480784': '김포간판 ok디자인 유리간판 제작과정',
    '223276935785': '인천간판 LED간판제작 LED간판가격 합리적이네요',
    '223231770726': '인천간판 실외간판 제작부터 시공까지',
    '223231770248': '인천간판 네온간판제작 빠짐없이 꼼꼼하게',
    '223250589647': '인천간판 천갈이 시공 잘하는 업체',
    '223285679116': '인천간판 ok디자인 간판의 종류와 과정',
    '223276936629': '유리창 LED간판 외부 간판 시공 과정'
}

# 키워드 리스트
keywords = ['kin', 'blog', 'daum', 'tistory', 'ad', 'search']

# 날짜별 키워드 중복 횟수를 저장할 딕셔너리
keyword_counts_by_date = {date_number: {keyword: 0 for keyword in keywords} for date_number in date_numbers}

# 각 날짜에 대해 데이터 수집
for date_number in date_numbers:
    # 날짜 형식이 잘못된 경우 처리
    if not re.match(r'^\d{4}$', date_number):
        print(f"잘못된 날짜 형식이기에 변경하여 처리합니다.: {date_number}")
    else:
        # 날짜 포맷을 변경 (예: "1023" -> "2024-10-23")
        year = "2024"
        month = date_number[:2]
        day = date_number[2:]
        date_input = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        all_dates.append(date_input)
        
    # "날짜" 옵션 선택
    select = Select(driver.find_element(By.ID, "sch_sort"))
    select.select_by_value("vi_date")

    # "날짜" 검색 필드에 값을 입력
    date_input_element = driver.find_element(By.ID, "sch_word")
    date_input_element.clear()
    date_input_element.send_keys(date_input)

    # 검색 버튼 클릭
    search_button = driver.find_element(By.CSS_SELECTOR, 'input[type="submit"].btn_submit')
    search_button.click()

    # WebDriverWait를 사용하여 테이블이 로드될 때까지 대기
    wait = WebDriverWait(driver, 10)
    table = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#container > div.tbl_wrap.tbl_head01 > table')))

    # 데이터를 저장할 리스트 초기화
    connection_path_data = []

    # 테이블 내의 데이터 추출
    if table:
        rows = table.find_elements(By.TAG_NAME, 'tr')
        for row in rows[1:]:  # 첫 번째 행은 헤더이므로 건너뜁니다.
            columns = row.find_elements(By.TAG_NAME, 'td')
            if len(columns) == 6:
                # "접속경로" 열만 추출하여 리스트에 추가
                connection_path_data.append(columns[1].text.strip())

    # "접속경로" 데이터를 이용하여 데이터프레임 생성
    data = {
        '접속 경로': connection_path_data
    }

    df = pd.DataFrame(data)

    # 각 DataFrame을 all_data에 추가
    all_data.append(df)
    # 키워드 중복 횟수 계산
    for keyword in keywords:
        count = df.applymap(lambda x: keyword in str(x).lower()).sum().sum()
        keyword_counts_by_date[date_number][keyword] += count

# 모든 데이터를 하나로 병합
combined_df = pd.concat(all_data, ignore_index=True)

# 데이터프레임 출력
print(combined_df)

# 데이터프레임 생성
keyword_counts_df = pd.DataFrame(keyword_counts_by_date).transpose()
keyword_counts_df = keyword_counts_df[keywords]  # 키워드 순서대로 정렬

# 인덱스를 날짜 포맷으로 변환하여 정렬
keyword_counts_df.index = pd.to_datetime(keyword_counts_df.index, format='%m%d')
keyword_counts_df = keyword_counts_df.sort_index()

# 인덱스를 'MM-DD' 형식으로 변경
keyword_counts_df.index = keyword_counts_df.index.strftime('%m-%d')

# 인덱스를 날짜로, 열을 키워드로 설정하여 데이터프레임 변환
keyword_counts_df = keyword_counts_df.transpose()
keyword_counts_df = keyword_counts_df.reset_index()
keyword_counts_df = keyword_counts_df.rename(columns={'index': ''})
print(keyword_counts_df)


# 한글만 추출하는 함수
def extract_korean(text):
    return ''.join(re.findall('[가-힣]+|[Ll][Ee][Dd](?=[^&]+)|(?<=query=)[^&]+', re.sub(r'[^\w\s]', '', text))) #query문만 가능 acq 무시가능성

# 제외하고 싶은 단어 리스트
exclude_words = ['', '접속경로', '브라우저','접속기기','일시','기타']  # <--- 원하는 단어들로 수정

# 데이터의 모든 셀에서 한글 추출하기
korean_texts = []
for column in combined_df.columns:
    for cell in combined_df[column]:
        if isinstance(cell, str):
            extracted_word = extract_korean(cell)
            if extracted_word not in exclude_words:  # 단어가 제외 목록에 없을 때만 추가
                korean_texts.append(extracted_word)

word_counts = Counter(korean_texts)

df2 = pd.DataFrame({'유입경로': list(word_counts.keys()), 'Count': list(word_counts.values())})

# "접속경로"에서 블로그 뒤의 숫자를 찾아 발생 횟수 계산
post_numbers = []
pattern1 = r'blog\.naver\.com\/[^\/]+\/(\d+)'  # 패턴 1
pattern2 = r'blog\.naver\.com\/PostView\.naver\?[^&]*&logNo=(\d+)'  # 패턴 2

for cell in combined_df['접속 경로']:
    match1 = re.search(pattern1, cell)
    match2 = re.search(pattern2, cell)
    if match1:
        post_numbers.append(match1.group(1))
    elif match2:
        post_numbers.append(match2.group(1))
    else:
        post_numbers.append(None)

# 발생 횟수 계산 및 데이터 저장
result_data = []
for post_number, count in Counter(post_numbers).items():
    if post_number is not None:  # None 값을 제거하여 처리
        title = data_dict.get(post_number, f'Unknown Post ({post_number})')  # data_dict에 없으면 "Unknown Post"로 표시
        result_data.append({'유입경로': title, 'Count': count})

df3 = pd.DataFrame(result_data)


# 발생 횟수 기준으로 정렬
if 'Count' in df3.columns:
    df_sorted3 = df3.sort_values(by='Count', ascending=False)
else:
    print("'Count' 열을 찾을 수 없습니다. df3 구조를 확인하십시오.")
    df_sorted3 = df3

# 데이터 프레임 출력
df_sorted = df2.sort_values(by='Count', ascending=False)
blank_row = pd.DataFrame({'유입경로': [''], 'Count': ['']})
new_row = pd.DataFrame({'유입경로': ['Keyword'], 'Count': ['Count']})
new_row2 = pd.DataFrame({'유입경로': ['제목'], 'Count': ['월보/스블 횟수']})
df_combined_final = pd.concat([keyword_counts_df, blank_row, new_row, df_sorted, blank_row, new_row2, df_sorted3], ignore_index=True)
print(df_combined_final)

# 데이터 Google Sheets에 업로드
upload_to_google_sheets(df_combined_final, sheet)
print("데이터가 Google Sheets에 성공적으로 업로드되었습니다.")