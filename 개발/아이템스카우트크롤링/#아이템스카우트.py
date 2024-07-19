import time
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re

# Google API 사용을 위한 인증 정보 설정
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file('credentials.json', scopes=scope)
client = gspread.authorize(creds)

# 구글 스프레드시트 및 시트 열기
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1ztBKMZJy0d7PJSDZH7vbXnv0ZLKT63AmjMZtK1_4tn4/edit#gid=0")
sheet = spreadsheet.worksheet('시트')

# 제품 설명 및 링크
product_info = {
    "클렌징폼": {"설명": "클렌징폼", "링크": "https://smartstore.naver.com/chicblanco/products/10087448986"},
    "크림": {"설명": "크림", "링크": "https://smartstore.naver.com/chicblanco/products/10087431491"},
    "미스트": {"설명": "미스트", "링크": "https://smartstore.naver.com/chicblanco/products/10087417416"},
    "팩트": {"설명": "팩트", "링크": "https://smartstore.naver.com/chicblanco/products/10087395963"},
    "미스트1": {"설명": "미스트1", "링크": "https://smartstore.naver.com/chicblanco/products/10380858914"},
    "크림1": {"설명": "크림1", "링크": "https://smartstore.naver.com/chicblanco/products/10380905131"},
    "선크림1": {"설명": "선크림1", "링크": "https://smartstore.naver.com/chicblanco/products/10380970867"}
}

# 로그인 함수 정의
def login(driver):
    # 로그인 페이지로 이동합니다.
    driver.get("https://itemscout.io/")

    try:
        # 로그인 버튼 클릭
        login_link = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#app > div > header > div > div.app-bar-right-content > a > span"))
        )
        login_link.click()

        # 이메일 입력란에 이메일 입력
        email_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='이메일']"))
        )
        email_input.send_keys("sbfrog@naver.com")

        # 비밀번호 입력란에 비밀번호 입력
        password_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='비밀번호']"))
        )
        password_input.send_keys("dbxodn0319-")

        # 로그인 버튼 클릭
        login_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '이메일 로그인')]/ancestor::button"))
        )
        login_button.click()

        # "일간" 버튼 클릭
        daily_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '일간')]"))
        )
        daily_button.click()

    except Exception as e:
        print("로그인에 실패했습니다. 다시 시도합니다.")
        print(e)
        driver.quit()  
        time.sleep(5)  
        driver = webdriver.Chrome()  
        login(driver)  

    return driver

# Selenium 웹 드라이버 초기화
driver = webdriver.Chrome()
driver = login(driver)

try:
    # 원하는 ID를 가진 상품 클릭
    product_ids = ["10087448986", "10087431491", "10087417416", "10087395963", "10380858914", "10380905131", "10380970867"]  
    start_row = 1  # 스프레드시트에 쓰기 시작할 행 번호
    for product_id in product_ids:
        product_element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, f"//div[@class='product-id'][contains(text(), '(ID {product_id})')]"))
        )
        driver.execute_script("arguments[0].click();", product_element)

        table_rows = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.XPATH, "//tr[@class='is-item-row tracking-one-product-tr']"))
        )

        df_data = []
        
        for row in table_rows:
            cells = row.find_elements(By.XPATH, ".//td")
            row_data = [cell.text.replace('\n', ' 변동') for cell in cells]  
            df_data.append(row_data)

        df = pd.DataFrame(df_data)

        table = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//table[@class='header-table max-width']"))
        )
        table_html = table.get_attribute('outerHTML')

        df2 = pd.read_html(table_html)[0]

        df.columns = df2.columns

        for col in df.columns:
            df[col] = df[col].astype(str)

        # 각 제품에 대한 설명 추가
        product_name = next((name for name, value in product_info.items() if product_id in value["링크"]), None)
        if product_name:  # Check if product_name is not None
            description = product_info[product_name]["설명"]
            link = product_info[product_name]["링크"]

            # 설명과 링크를 별도의 행으로 추가
            description_link_df = pd.DataFrame({"Description": [f"{description}, {link}"]})
            merged_df = pd.concat([description_link_df, df2, df], ignore_index=True)

            # 변동+숫자를 없애는 작업
            for col in merged_df.columns[1:]:
                merged_df[col] = merged_df[col].str.replace(r'\s변동\d+', '', regex=True)
                merged_df[col] = merged_df[col].str.replace(r'\,\d+', '', regex=True)
                merged_df[col] = merged_df[col].str.replace(r'\d+\위\s+-', '', regex=True)  # 숫자, '위', 공백, '-' 패턴 제거
                merged_df[col] = merged_df[col].str.replace(r'(\d+P\s+\d+)', r'\g<1>위', regex=True)

            # 스프레드시트에 데이터프레임 쓰기
            set_with_dataframe(sheet, merged_df, row=start_row, col=1, include_column_header=True, include_index=False)
            start_row += len(merged_df) + 2  # 다음 데이터프레임을 쓸 행 번호 업데이트

        driver.back()

finally:
    driver.quit()
