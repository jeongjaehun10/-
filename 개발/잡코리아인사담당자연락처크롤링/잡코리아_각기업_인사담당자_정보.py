import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup as bs
import pandas as pd
import time
import datetime
import threading

class JobKoreaScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("JobKorea 데이터 수집 프로그램")

        self.made_by = tk.Label(self.root, text="Made by 기획홍보파트", font=("Helvetica", 12, "italic"))
        self.made_by.pack(pady=10)

        self.page_limit_notice = tk.Label(self.root, text="페이지범위는 10페이지 이내", font=("Helvetica", 12, "italic"))
        self.page_limit_notice.pack(pady=5)

        self.login_notice = tk.Label(self.root, text="반드시 로그인 후 확인", font=("Helvetica", 12, "italic"))
        self.login_notice.pack(pady=5)

        self.label_keyword = tk.Label(self.root, text="업무 키워드:")
        self.label_keyword.pack(pady=5)
        self.entry_keyword = tk.Entry(self.root, width=30)
        self.entry_keyword.pack()

        self.label_start_page = tk.Label(self.root, text="시작 페이지:")
        self.label_start_page.pack(pady=5)
        self.entry_start_page = tk.Entry(self.root, width=10)
        self.entry_start_page.pack()

        self.label_end_page = tk.Label(self.root, text="종료 페이지:")
        self.label_end_page.pack(pady=5)
        self.entry_end_page = tk.Entry(self.root, width=10)
        self.entry_end_page.pack()

        self.btn_run = tk.Button(self.root, text="시작", command=self.open_login_page)
        self.btn_run.pack(pady=10)

        self.btn_stop = tk.Button(self.root, text="종료", command=self.stop_scraping)
        self.btn_stop.pack(pady=5)

        self.status_label = tk.Label(self.root, text="")
        self.status_label.pack(pady=10)

        self.scraping = False

    def open_login_page(self):
        options = Options()
        # options.add_argument("--headless")  # 브라우저를 숨겨서 실행할 때 사용
        self.driver = webdriver.Chrome(options=options)

        login_url = "https://www.jobkorea.co.kr/Login/Login_Tot.asp?rDBName=GG&re_url=/"
        self.driver.get(login_url)

        messagebox.showinfo("로그인 안내", "로그인을 완료한 후 '확인'을 눌러주세요.")
        
        self.btn_run.config(text="확인(팝업창 확인 눌러야 시작)", command=self.start_scraping)

    def start_scraping(self):
        self.scraping = True
        threading.Thread(target=self.scrape_jobs).start()

    def scrape_jobs(self):
        keyword = self.entry_keyword.get()
        start_page = self.entry_start_page.get()
        end_page = self.entry_end_page.get()

        if not keyword:
            messagebox.showerror("입력 오류", "검색 키워드를 입력해주세요.")
            return

        if not start_page.isdigit() or not end_page.isdigit():
            messagebox.showerror("입력 오류", "시작 페이지와 종료 페이지를 올바르게 입력해주세요.")
            return

        start_page = int(start_page)
        end_page = int(end_page)

        if start_page <= 0 or end_page < start_page:
            messagebox.showerror("입력 오류", "올바른 페이지 범위를 입력해주세요.")
            return

        try:
            search_url = f"https://www.jobkorea.co.kr/Search/?stext={keyword}&tabType=recruit&Page_No=1"
            self.driver.get(search_url)
            time.sleep(3)  # 페이지 로드를 기다림

            soup = bs(self.driver.page_source, 'html.parser')

            total_jobs = soup.find('p', class_='filter-text').find('strong').text.replace(',', '')
            total_pages = (int(total_jobs) + 19) // 20

            if end_page > total_pages:
                end_page = total_pages

            # Display total jobs and pages information without requiring user confirmation
            messagebox.showinfo("데이터 수집 시작", f"총 {total_jobs}건의 검색 결과가 있습니다.\n총 {total_pages} 페이지 중\n{start_page} 페이지부터 {end_page} 페이지까지 수집합니다.")

            df = pd.DataFrame(columns=['회사명', '공고', '직장위치', '인사 담당자', '연락처', '이메일'])

            collected_count = 0
            for page in range(start_page, end_page + 1):
                if not self.scraping:
                    break

                search_url = f"https://www.jobkorea.co.kr/Search/?stext={keyword}&tabType=recruit&Page_No={page}"
                self.driver.get(search_url)
                time.sleep(3)

                soup = bs(self.driver.page_source, 'html.parser')

                company_names = soup.select('a.name.dev_view')
                info = soup.select('a.title.dev_view')
                locations = soup.select('span.loc.long')

                job_links = []

                for j in range(len(info)):
                    try:
                        job_url = info[j]['href']
                        job_links.append("https://www.jobkorea.co.kr" + job_url)
                    except IndexError as e:
                        print(f"IndexError: Could not find job link for index {j}")
                    except Exception as e:
                        print(f"Error occurred while extracting job link: {e}")

                for i, job_url in enumerate(job_links):
                    if not self.scraping:
                        break

                    try:
                        self.driver.get(job_url)
                        time.sleep(2)

                        soup_detail = bs(self.driver.page_source, 'html.parser')

                        company = company_names[i].text.strip() if i < len(company_names) else ''
                        job_title = info[i].text.strip() if i < len(info) else ''
                        
                        workplace = locations[i].text.strip() if i < len(locations) else ''

                        try:
                            for _ in range(10):
                                self.driver.execute_script("window.scrollBy(0, 3000);")
                                time.sleep(1)
                                soup_detail = bs(self.driver.page_source, 'html.parser')
                                if soup_detail.select_one('div.manager'):
                                    break
                        except Exception as e:
                            print(f"스크롤 중 오류 발생: {e}")

                        hr_manager = ''
                        email = ''
                        contact = ''

                        manager_section = soup_detail.select_one('div.manager')
                        if manager_section:
                            hr_manager_tag = manager_section.select_one('dt:contains("인사 담당자")')
                            if hr_manager_tag:
                                hr_manager = hr_manager_tag.find_next('dd').text.strip()

                            contact_element = manager_section.select_one('dd.devTplLyClick span.tahoma.tplHide')
                            if contact_element:
                                contact = contact_element.text.strip()

                            email_element = manager_section.select_one('a.devChargeEmail')
                            if email_element:
                                email = email_element.text.strip()

                        df.loc[len(df)] = [company, job_title, workplace, hr_manager, contact, email]
                        collected_count += 1

                    except Exception as ex:
                        print(f"Error occurred while scraping job details: {ex}")

            file_name = f"{keyword}_잡코리아_기업인사담당자_정보.csv"
            df.to_csv(file_name, index=False, encoding='utf-8-sig')
            self.status_label.config(text=f"CSV 파일 저장 완료: {file_name}")

        finally:
            self.driver.quit()

    def stop_scraping(self):
        self.scraping = False

if __name__ == "__main__":
    root = tk.Tk()
    app = JobKoreaScraperApp(root)
    root.mainloop()
