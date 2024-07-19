import subprocess
import sys
from importlib import util

# 필요한 패키지를 확인하고 없는 패키지 목록을 반환하는 함수
def check_required_packages(required_packages):
    missing_packages = []
    for package in required_packages:
        if util.find_spec(package) is None:
            missing_packages.append(package)
    return missing_packages

# 필요한 패키지 목록
required_packages = ['pandas', 'selenium', 'openpyxl']

# 필요한 패키지 중에서 설치되지 않은 패키지 확인
missing_packages = check_required_packages(required_packages)

# 필요한 패키지가 설치되지 않은 경우, 자동으로 설치
if missing_packages:
    print("필요한 패키지가 설치되지 않았습니다. 라이브러리를 자동으로 설치합니다.")
    for package in missing_packages:
        subprocess.call([sys.executable, "-m", "pip", "install", package])
    print("라이브러리 설치가 완료되었습니다.")

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import re
from collections import Counter
from selenium.webdriver.chrome.options import Options


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


# 여러 날짜를 입력받아 리스트로 분할
date_numbers = input("검색할 날짜를 입력하세요(하루 입력:1025, 여러 날짜:1023,1024,1025): ").split(',')
print("약 30초 소요됩니다.")

# 모든 데이터를 저장할 리스트 초기화
all_data = []
all_dates = []
# 각 날짜에 대해 데이터 수집
for date_number in date_numbers:
    # 날짜 형식이 잘못된 경우 처리
    if not re.match(r'^\d{4}$', date_number):
        print(f"잘못된 날짜 형식이기에 변경하여 처리합니다.: {date_number}")
    else:
        # 날짜 포맷을 변경 (예: "1023" -> "2023-10-23")
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
    ip_data = []
    connection_path_data = []
    browser_data = []
    os_data = []
    device_data = []
    date_data = []

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

# 모든 데이터를 하나로 병합
combined_df = pd.concat(all_data, ignore_index=True)

# 데이터프레임 출력
print(combined_df)


# 키워드 리스트
keywords = ['kin', 'blog', 'daum', 'tistory', 'ad', 'search']

# 각 키워드에 대한 중복 횟수 확인 후 데이터프레임 생성
keyword_counts = []

for keyword in keywords:
    count = combined_df.applymap(lambda x: keyword in str(x).lower()).sum().sum()
    keyword_counts.append(count)

# 데이터프레임 생성
keyword_counts_df = pd.DataFrame({'유입경로': keywords, 'Count': keyword_counts})

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

# 데이터 프레임 출력
df_sorted = df2.sort_values(by='Count', ascending=False)
blank_row = pd.DataFrame({'유입경로': [''], 'Count': ['']})
new_row = pd.DataFrame({'유입경로': ['Keyword'], 'Count': ['Count']})
df_combined_final = pd.concat([keyword_counts_df,blank_row,new_row,df_sorted], ignore_index=True)
print(df_combined_final)
# Save the combined DataFrame to an Excel file
file_name = ', '.join(all_dates) + " 간판 접속 키워드.xlsx"
df_combined_final.to_excel(file_name, index=False)
print(f"{file_name}이 성공적으로 저장되었습니다.")

