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

# 필요한 패키지 불러오기
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import filedialog

# 웹 드라이버 초기화
options = Options()
#options.add_argument("--headless")  # 웹 브라우저 화면을 표시하지 않음

# tk 윈도우 생성
root = tk.Tk()
root.withdraw()  # 기본 창 숨기기

# 파일 대화상자 열기
excel_file_path = filedialog.askopenfilename()

# 선택된 파일 경로 확인
if not excel_file_path:
    print("파일 선택이 취소되었습니다.")
    sys.exit()
    
def extract_cafe_urls_from_excel(file_path, sheet_name, url_column_name, cafe_name_column, member_count_column):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cafe_urls = df[url_column_name].tolist()
    cafe_names = df[cafe_name_column].tolist()
    if member_count_column:
        member_counts = df[member_count_column].tolist()
    else:
        member_counts = [None] * len(df)
    return cafe_urls, cafe_names, member_counts

# 사용자로부터 엑셀 파일 경로, 시트 이름, 열 이름 입력 받기
sheet_name = input("시트 이름을 입력하세요(예시: Sheet1): ")
url_column_name = input("URL이 들어있는 열 이름을 입력하세요(예시: 링크): ")
cafe_name_column = input("카페명이 들어있는 열 이름을 입력하세요(예시: 카페명): ")
member_count_column = input("회원 수가 들어있는 열 이름을 입력하세요 (없으면 엔터): ")
output_filename = input("저장할 파일 이름을 입력하세요 (확장자 없이, 중복x): ")
print("저장위치는 local 영역인 user에 있음")

cafe_urls, cafe_names, member_counts = extract_cafe_urls_from_excel(excel_file_path, sheet_name, url_column_name, cafe_name_column, member_count_column)

# 사용자로부터 키워드 입력 받기
keywords = input("키워드 목록을 입력하세요 (쉼표로 구분): ").split(",")

# 결과를 저장할 데이터프레임 초기화
df = pd.DataFrame(columns=['카페명', '카페 URL', '회원 수'] + keywords + ['총합'])

for cafe_url, cafe_name, member_count in zip(cafe_urls, cafe_names, member_counts):
    try:
        driver = webdriver.Chrome(options=options)
        driver.get(cafe_url)  # 입력된 URL을 받아오기
    except:
        print(f"{cafe_url}은(는) 잘못된 URL입니다.")
        if 'driver' in locals():
            driver.quit()
        # "오류"를 총합에 추가
        data = [cafe_name, cafe_url, member_count] + [0] * len(keywords) + ["오류"]
        df.loc[len(df)] = data
        continue  # 잘못된 URL인 경우 다음 URL로 넘어가도록 처리

    # 키워드 별로 작업 수행
    keyword_counts = {keyword: 0 for keyword in keywords}
    error_occurred = False  # 오류 여부를 나타내는 플래그
    for keyword in keywords:
        try:
            # 'query' 이름을 가진 검색 상자 요소를 찾을 때 대기
            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "query"))
            )
            search_box.clear()
            search_box.send_keys(keyword)
            search_box.send_keys(Keys.RETURN)

            # 'cafe_main' 프레임을 찾을 때 대기
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, 'cafe_main'))
            )

            # 'cafe_main' 프레임 안에서 작업을 수행
            article_elements = driver.find_elements(By.CSS_SELECTOR, 'td.td_article')
            if len(article_elements) > 0:
                keyword_counts[keyword] = 1
        except:
            print(f"{cafe_url}에 대한 작업 중에 문제가 발생했습니다.")
            error_occurred = True  # 오류가 발생한 경우 플래그 설정
            break  # 키워드에 대한 작업 중 오류 발생 시 해당 URL로부터 빠져나감

        # 메인 프레임으로 다시 전환
        driver.switch_to.default_content()

    # 카페별 키워드별 카운트 합계 계산
    if error_occurred:
        total_counts = "오류"  # 오류가 발생한 경우 "오류"로 표시
    else:
        total_counts = sum(keyword_counts.values())

    # 결과를 데이터프레임에 추가
    data = [cafe_name, cafe_url, member_count] + list(keyword_counts.values()) + [total_counts]
    df.loc[len(df)] = data

    # 웹 드라이버 종료
    driver.quit()

# 데이터프레임 출력
print(df)

# 데이터프레임을 엑셀 파일로 내보내기
df.to_excel(output_filename + '.xlsx', index=False)

