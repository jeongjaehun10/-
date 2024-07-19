import subprocess
import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

# 웹 드라이버 초기화
options = Options()
# options.add_argument("--headless")  # 웹 브라우저 화면을 표시하지 않음

# tk 윈도우 생성
root = tk.Tk()
root.withdraw()  # 기본 창 숨기기

# 파일 대화상자 열기
excel_file_path = filedialog.askopenfilename()

# 선택된 파일 경로 확인
if not excel_file_path:
    print("파일 선택이 취소되었습니다.")
    sys.exit()

# 사용자로부터 엑셀 파일 경로, 시트 이름, 열 이름 입력 받기
sheet_name = input("시트 이름을 입력하세요(예시: Sheet1): ")
url_column_name = input("URL이 들어있는 열 이름을 입력하세요(예시: 링크): ")
cafe_name_column = input("카페명이 들어있는 열 이름을 입력하세요(예시: 카페명): ")
member_count_column = input("회원 수가 들어있는 열 이름을 입력하세요 (없으면 엔터): ")
output_filename = input("저장할 파일 이름을 입력하세요 (확장자 없이, 중복x): ")
print("저장위치는 실행파일")

# 엑셀 파일에서 카페 URL 및 관련 정보 추출
def extract_cafe_urls_from_excel(file_path, sheet_name, url_column_name, cafe_name_column, member_count_column):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cafe_urls = df[url_column_name].tolist()
    cafe_names = df[cafe_name_column].tolist()
    if member_count_column:
        member_counts = df[member_count_column].tolist()
    else:
        member_counts = [None] * len(df)
    return cafe_urls, cafe_names, member_counts

cafe_urls, cafe_names, member_counts = extract_cafe_urls_from_excel(excel_file_path, sheet_name, url_column_name, cafe_name_column, member_count_column)

# 사용자로부터 키워드 입력 받기
keywords = input("키워드 목록을 입력하세요 (쉼표로 구분): ").split(",")
print("페이지 당 게시글 수는 15건입니다. 단, 마지막 페이지는 15건이 안될 수도 있습니다.")
def extract_cafe_info(driver, cafe_url, cafe_name, member_count, keywords):
    keyword_counts = {}
    keyword_max_pages = {}  # Dictionary to store max pages for each keyword
    total_counts = 0

    for keyword in keywords:
        try:
            driver.get(cafe_url)
            search_box = driver.find_element(By.NAME, "query")
            search_box.clear()
            search_box.send_keys(keyword)
            search_box.send_keys(Keys.RETURN)
            time.sleep(2)  # Wait for the search results to load
            driver.switch_to.frame("cafe_main")

            while True:
                # 특정 CSS 선택자를 사용하여 요소를 확인하고 카운트 증가
                article_elements = driver.find_elements(By.CSS_SELECTOR, 'td.td_article')
                page_elements = driver.find_element(By.CSS_SELECTOR, 'div.prev-next')
                page_string = re.sub('[^0-9 ]', '', page_elements.text)
                pages = [item for item in page_string.split(' ') if item != '']

                if pages:
                    new_pages = [int(item) for item in pages]
                    items = [item for item in article_elements]
                    total_counts += len(items) * max(new_pages)
                    keyword_counts[keyword] = len(items) * max(new_pages)

                    # Store the max pages for the keyword
                    keyword_max_pages[keyword] = max(new_pages)

                    # Click the "다음" button to go to the next page
                    try:
                        next_page_button = driver.find_element(By.CSS_SELECTOR, '#main-area > div.prev-next > a.pgR')
                        next_page_button.click()
                        time.sleep(2)  # Wait for the next page to load
                    except NoSuchElementException:
                        # No more next pages, exit the loop
                        break
                else:
                    # No more pages, exit the loop
                    break
        except Exception as e:
            print(f"An error occurred while processing {cafe_url} for the keyword '{keyword}': {str(e)}")

    return cafe_name, cafe_url, member_count, total_counts, keyword_counts, keyword_max_pages

# ... (Previous code remains unchanged)

# 결과를 저장할 데이터프레임 초기화
columns = ['카페명', '카페 URL', '회원 수'] + [f'{keyword}_총페이지' for keyword in keywords] + ['총합']
df = pd.DataFrame(columns=columns)

# 각 카페별로 작업 수행
for cafe_url, cafe_name, member_count in zip(cafe_urls, cafe_names, member_counts):
    try:
        driver = webdriver.Chrome(options=options)
        result = extract_cafe_info(driver, cafe_url, cafe_name, member_count, keywords)

        # Create a list to hold the data for the current cafe
        cafe_data = [result[0], result[1], result[2]]

        # Add the total pages for each keyword to the list
        for keyword in keywords:
            cafe_data.append(result[5].get(keyword, 0))

        # Add the overall sum of pages for all keywords to the list
        total_pages = sum(result[5].values())
        cafe_data.append(total_pages)

        # Append the list as a new row to the main DataFrame
        df.loc[len(df)] = cafe_data
    except Exception as e:
        print(f"An error occurred while processing {cafe_url}: {str(e)}")
    finally:
        driver.quit()  # Always quit the driver to avoid resource leaks

# 데이터프레임 출력
print(df)

# 데이터프레임을 엑셀 파일로 내보내기
df.to_excel(output_filename + '.xlsx', index=False)
