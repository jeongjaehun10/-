from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd
import re

def crawl_cafe_info(url):
    # Headless 모드 설정
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')

    # 크롬 드라이버 생성
    driver = webdriver.Chrome(options=chrome_options)

    # 크롤링할 카페 URL
    driver.get(url)

    # 카페명 가져오기
    try:
        h1_tag = driver.find_element('css selector', 'h1.d-none')
        cafe_name = h1_tag.text
    except Exception as e:
        cafe_name = '카페명을 찾을 수 없습니다.'

    # cafe_main으로 전환
    cafe_main_url = url + '?iframe_url=/cafeMain.nhn%3Fclubid=12345678'  # 12345678은 실제 카페의 clubid로 변경
    driver.get(cafe_main_url)

    # 원하는 작업 수행...

    # 회원수 가져오기
    try:
        em_tag = driver.find_element('css selector', '#cafe-info-data div.box-g div.ia-info-data2 ul li.mem-cnt-info a:nth-child(2) em')
        member_count_str = em_tag.text
        member_count = int(re.sub('[^0-9]', '', member_count_str))
    except Exception as e:
        member_count = 0

    # 드라이버 종료
    driver.quit()

    return {'카페명': cafe_name, 'URL': url, '회원수': member_count}

# 여러 개의 URL을 ','로 구분하여 입력
url_input = input("여러 개의 카페 URL을 ','로 구분하여 입력하세요: ")
urls = url_input.split(',')

# 각 카페에 대한 정보를 담을 리스트
cafe_data_list = []

# 각 URL에 대해 crawl_cafe_info 함수 호출
for url in urls:
    cafe_data_list.append(crawl_cafe_info(url))

# 결과 데이터프레임 생성
result_df = pd.DataFrame(cafe_data_list)

# 결과 출력
print(result_df)

# 결과를 엑셀 파일로 저장
excel_filename = 'cafe_info.xlsx'
result_df.to_excel(excel_filename, index=False)
print(f'데이터프레임이 {excel_filename}로 저장되었습니다.')
