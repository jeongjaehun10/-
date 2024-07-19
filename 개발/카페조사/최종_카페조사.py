import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import threading
import re
from selenium.common.exceptions import NoSuchElementException

# 전역 변수
stop_event = threading.Event()
progress_popup = None
progress_label = None
progress_text = None
total_items = 0
current_item = 0

def open_file_dialog(entry):
    file_path = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def extract_cafe_urls_from_excel(file_path, sheet_name, url_column_name, cafe_name_column, member_count_column):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    cafe_urls = df[url_column_name].tolist()
    cafe_names = df[cafe_name_column].tolist()
    if member_count_column:
        member_counts = df[member_count_column].tolist()
    else:
        member_counts = [None] * len(df)
    return cafe_urls, cafe_names, member_counts

def update_progress(message, current, total):
    global progress_text
    if progress_label:
        progress_label.config(text=message)
        progress_text.config(text=f'총 {total}개의 카페 중 {current}번 카페 처리 중')
        progress_popup.update_idletasks()

def show_progress_popup():
    global progress_popup, progress_label, progress_text
    progress_popup = tk.Toplevel()
    progress_popup.title("진행 중")
    progress_label = tk.Label(progress_popup, text="작업이 진행 중입니다...")
    progress_label.pack(pady=10, padx=10)
    progress_text = tk.Label(progress_popup, text="총 0개의 카페 중 0번 카페 처리 중")
    progress_text.pack(pady=10, padx=10)
    button_stop = tk.Button(progress_popup, text="중지", command=stop_processing)
    button_stop.pack(pady=10)
    progress_popup.transient(root)
    progress_popup.grab_set()

def process_excel(entry_file_path, entry_sheet_name, entry_url_column_name, entry_cafe_name_column, entry_member_count_column, entry_keywords, entry_output_filename):
    global stop_event, total_items, current_item
    stop_event.clear()

    excel_file_path = entry_file_path.get()
    sheet_name = entry_sheet_name.get()
    url_column_name = entry_url_column_name.get()
    cafe_name_column = entry_cafe_name_column.get()
    member_count_column = entry_member_count_column.get()
    output_filename = entry_output_filename.get()

    cafe_urls, cafe_names, member_counts = extract_cafe_urls_from_excel(excel_file_path, sheet_name, url_column_name, cafe_name_column, member_count_column)
    keywords = entry_keywords.get().split(",")
    df = pd.DataFrame(columns=['카페명', '카페 URL', '회원 수'] + keywords + ['총합'])

    total_items = len(cafe_urls)
    current_item = 0

    options = Options()
    options.add_argument("--headless")

    for cafe_url, cafe_name, member_count in zip(cafe_urls, cafe_names, member_counts):
        if stop_event.is_set():
            df.to_excel(output_filename + '_중지됨.xlsx', index=False)
            messagebox.showinfo("중지됨", "작업이 중지되었습니다.")
            return

        current_item += 1
        update_progress(f"{cafe_name} 처리 중...", current_item, total_items)

        try:
            driver = webdriver.Chrome(options=options)
            driver.get(cafe_url)
        except:
            print(f"{cafe_url}은(는) 잘못된 URL입니다.")
            if 'driver' in locals():
                driver.quit()
            data = [cafe_name, cafe_url, member_count] + [0] * len(keywords) + ["오류"]
            df.loc[len(df)] = data
            continue

        keyword_counts = {keyword: 0 for keyword in keywords}
        error_occurred = False
        for keyword in keywords:
            try:
                search_box = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "query"))
                )
                search_box.clear()
                search_box.send_keys(keyword)
                search_box.send_keys(Keys.RETURN)
                WebDriverWait(driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.NAME, 'cafe_main'))
                )
                article_elements = driver.find_elements(By.CSS_SELECTOR, 'td.td_article')
                if len(article_elements) > 0:
                    keyword_counts[keyword] = 1
            except:
                print(f"{cafe_url}에 대한 작업 중에 문제가 발생했습니다.")
                error_occurred = True
                break
            driver.switch_to.default_content()

        if error_occurred:
            total_counts = "오류"
        else:
            total_counts = sum(keyword_counts.values())

        data = [cafe_name, cafe_url, member_count] + list(keyword_counts.values()) + [total_counts]
        df.loc[len(df)] = data
        driver.quit()

    df.to_excel(output_filename + '.xlsx', index=False)
    progress_popup.destroy()
    messagebox.showinfo("완료", "작업이 완료되었습니다.")

def process_excel2(entry_file_path_tab2, entry_sheet_name_tab2, entry_url_column_name_tab2, entry_cafe_name_column_tab2, entry_member_count_column_tab2, entry_keywords_tab2, entry_output_filename_tab2):
    global stop_event, total_items, current_item
    stop_event.clear()

    options = Options()
    options.add_argument("--headless")
    excel_file_path = entry_file_path_tab2.get()
    url_column_name = entry_url_column_name_tab2.get()
    cafe_name_column = entry_cafe_name_column_tab2.get()
    member_count_column = entry_member_count_column_tab2.get()
    output_filename = entry_output_filename_tab2.get()
    sheet_name = entry_sheet_name_tab2.get()
    keywords = entry_keywords_tab2.get().split(",")
    keywords = [keyword.strip() for keyword in keywords if keyword.strip()]

    def extract_cafe_info(driver, cafe_url, cafe_name, member_count, keywords):
        keyword_counts = {}
        keyword_max_pages = {}
        total_counts = 0

        for keyword in keywords:
            try:
                driver.get(cafe_url)
                search_box = driver.find_element(By.NAME, "query")
                search_box.clear()
                search_box.send_keys(keyword)
                search_box.send_keys(Keys.RETURN)
                time.sleep(2)
                driver.switch_to.frame("cafe_main")

                while True:
                    article_elements = driver.find_elements(By.CSS_SELECTOR, 'td.td_article')
                    page_elements = driver.find_element(By.CSS_SELECTOR, 'div.prev-next')
                    page_string = re.sub('[^0-9 ]', '', page_elements.text)
                    pages = [item for item in page_string.split(' ') if item != '']

                    if pages:
                        new_pages = [int(item) for item in pages]
                        items = [item for item in article_elements]
                        total_counts += len(items) * max(new_pages)
                        keyword_counts[keyword] = len(items) * max(new_pages)
                        keyword_max_pages[keyword] = max(new_pages)
                        try:
                            next_page_button = driver.find_element(By.CSS_SELECTOR, '#main-area > div.prev-next > a.pgR')
                            next_page_button.click()
                            time.sleep(2)
                        except NoSuchElementException:
                            break
                    else:
                        break
            except Exception as e:
                print(f"An error occurred while processing {cafe_url} for the keyword '{keyword}': {str(e)}")

        return cafe_name, cafe_url, member_count, total_counts, keyword_counts, keyword_max_pages

    columns = ['카페명', '카페 URL', '회원 수'] + [f'{keyword}_총페이지' for keyword in keywords] + ['총합']
    df = pd.DataFrame(columns=columns)
    cafe_urls, cafe_names, member_counts = extract_cafe_urls_from_excel(excel_file_path, sheet_name, url_column_name, cafe_name_column, member_count_column)

    total_items = len(cafe_urls)
    current_item = 0

    for cafe_url, cafe_name, member_count in zip(cafe_urls, cafe_names, member_counts):
        if stop_event.is_set():
            df.to_excel(output_filename + '_중지됨.xlsx', index=False)
            messagebox.showinfo("중지됨", "작업이 중지되었습니다.")
            return

        current_item += 1
        update_progress(f"{cafe_name} 처리 중...", current_item, total_items)

        try:
            driver = webdriver.Chrome(options=options)
            result = extract_cafe_info(driver, cafe_url, cafe_name, member_count, keywords)
            cafe_data = [result[0], result[1], result[2]]
            for keyword in keywords:
                cafe_data.append(result[5].get(keyword, 0))
            total_pages = sum(result[5].values())
            cafe_data.append(total_pages)
            df.loc[len(df)] = cafe_data
        except Exception as e:
            print(f"An error occurred while processing {cafe_url}: {str(e)}")
        finally:
            driver.quit()

    df.to_excel(output_filename + '.xlsx', index=False)
    progress_popup.destroy()
    messagebox.showinfo("완료", "작업이 완료되었습니다.")

def start_processing(entry_file_path, entry_sheet_name, entry_url_column_name, entry_cafe_name_column, entry_member_count_column, entry_keywords, entry_output_filename):
    global stop_event
    stop_event.clear()
    show_progress_popup()
    threading.Thread(target=process_excel, args=(entry_file_path, entry_sheet_name, entry_url_column_name, entry_cafe_name_column, entry_member_count_column, entry_keywords, entry_output_filename)).start()

def start_processing2(entry_file_path_tab2, entry_sheet_name_tab2, entry_url_column_name_tab2, entry_cafe_name_column_tab2, entry_member_count_column_tab2, entry_keywords_tab2, entry_output_filename_tab2):
    global stop_event
    stop_event.clear()
    show_progress_popup()
    threading.Thread(target=process_excel2, args=(entry_file_path_tab2, entry_sheet_name_tab2, entry_url_column_name_tab2, entry_cafe_name_column_tab2, entry_member_count_column_tab2, entry_keywords_tab2, entry_output_filename_tab2)).start()

def stop_processing():
    stop_event.set()

root = tk.Tk()
root.title("카페 정보 수집기")

# Tab 1
tab_control = ttk.Notebook(root)
tab1 = ttk.Frame(tab_control)
tab_control.add(tab1, text='카페키워드유무')

label_made_by_tab1 = tk.Label(tab1, text="Made by 기획홍보파트", font=("Helvetica", 12, "italic"))
label_made_by_tab1.grid(row=0, column=0, columnspan=3, pady=5)

label_file_path = tk.Label(tab1, text="엑셀 파일 경로:")
label_file_path.grid(row=1, column=0, padx=10, pady=10, sticky='e')
entry_file_path = tk.Entry(tab1, width=50)
entry_file_path.grid(row=1, column=1, padx=10, pady=10)
button_browse = tk.Button(tab1, text="찾아보기", command=lambda: open_file_dialog(entry_file_path))
button_browse.grid(row=1, column=2, padx=10, pady=10)

label_sheet_name = tk.Label(tab1, text="시트 이름:")
label_sheet_name.grid(row=2, column=0, padx=10, pady=10, sticky='e')
entry_sheet_name = tk.Entry(tab1, width=50)
entry_sheet_name.grid(row=2, column=1, padx=10, pady=10)

label_url_column_name = tk.Label(tab1, text="URL 열 이름:")
label_url_column_name.grid(row=3, column=0, padx=10, pady=10, sticky='e')
entry_url_column_name = tk.Entry(tab1, width=50)
entry_url_column_name.grid(row=3, column=1, padx=10, pady=10)

label_cafe_name_column = tk.Label(tab1, text="카페명 열 이름:")
label_cafe_name_column.grid(row=4, column=0, padx=10, pady=10, sticky='e')
entry_cafe_name_column = tk.Entry(tab1, width=50)
entry_cafe_name_column.grid(row=4, column=1, padx=10, pady=10)

label_member_count_column = tk.Label(tab1, text="회원 수 열 이름:")
label_member_count_column.grid(row=5, column=0, padx=10, pady=10, sticky='e')
entry_member_count_column = tk.Entry(tab1, width=50)
entry_member_count_column.grid(row=5, column=1, padx=10, pady=10)

label_keywords = tk.Label(tab1, text="검색 키워드 (콤마로 구분):")
label_keywords.grid(row=6, column=0, padx=10, pady=10, sticky='e')
entry_keywords = tk.Entry(tab1, width=50)
entry_keywords.grid(row=6, column=1, padx=10, pady=10)

label_output_filename = tk.Label(tab1, text="출력 파일명:")
label_output_filename.grid(row=7, column=0, padx=10, pady=10, sticky='e')
entry_output_filename = tk.Entry(tab1, width=50)
entry_output_filename.grid(row=7, column=1, padx=10, pady=10)

button_start = tk.Button(tab1, text="작업 시작", command=lambda: start_processing(entry_file_path, entry_sheet_name, entry_url_column_name, entry_cafe_name_column, entry_member_count_column, entry_keywords, entry_output_filename))
button_start.grid(row=8, column=1, padx=10, pady=10)

# Tab 2
tab2 = ttk.Frame(tab_control)
tab_control.add(tab2, text='키워드별페이지수')

label_made_by_tab2 = tk.Label(tab2, text="Made by 기획홍보파트", font=("Helvetica", 12, "italic"))
label_made_by_tab2.grid(row=0, column=0, columnspan=3, pady=5)

label_file_path_tab2 = tk.Label(tab2, text="엑셀 파일 경로:")
label_file_path_tab2.grid(row=1, column=0, padx=10, pady=10, sticky='e')
entry_file_path_tab2 = tk.Entry(tab2, width=50)
entry_file_path_tab2.grid(row=1, column=1, padx=10, pady=10)
button_browse_tab2 = tk.Button(tab2, text="찾아보기", command=lambda: open_file_dialog(entry_file_path_tab2))
button_browse_tab2.grid(row=1, column=2, padx=10, pady=10)

label_sheet_name_tab2 = tk.Label(tab2, text="시트 이름:")
label_sheet_name_tab2.grid(row=2, column=0, padx=10, pady=10, sticky='e')
entry_sheet_name_tab2 = tk.Entry(tab2, width=50)
entry_sheet_name_tab2.grid(row=2, column=1, padx=10, pady=10)

label_url_column_name_tab2 = tk.Label(tab2, text="URL 열 이름:")
label_url_column_name_tab2.grid(row=3, column=0, padx=10, pady=10, sticky='e')
entry_url_column_name_tab2 = tk.Entry(tab2, width=50)
entry_url_column_name_tab2.grid(row=3, column=1, padx=10, pady=10)

label_cafe_name_column_tab2 = tk.Label(tab2, text="카페명 열 이름:")
label_cafe_name_column_tab2.grid(row=4, column=0, padx=10, pady=10, sticky='e')
entry_cafe_name_column_tab2 = tk.Entry(tab2, width=50)
entry_cafe_name_column_tab2.grid(row=4, column=1, padx=10, pady=10)

label_member_count_column_tab2 = tk.Label(tab2, text="회원 수 열 이름:")
label_member_count_column_tab2.grid(row=5, column=0, padx=10, pady=10, sticky='e')
entry_member_count_column_tab2 = tk.Entry(tab2, width=50)
entry_member_count_column_tab2.grid(row=5, column=1, padx=10, pady=10)

label_keywords_tab2 = tk.Label(tab2, text="검색 키워드 (콤마로 구분):")
label_keywords_tab2.grid(row=6, column=0, padx=10, pady=10, sticky='e')
entry_keywords_tab2 = tk.Entry(tab2, width=50)
entry_keywords_tab2.grid(row=6, column=1, padx=10, pady=10)

label_output_filename_tab2 = tk.Label(tab2, text="출력 파일명:")
label_output_filename_tab2.grid(row=7, column=0, padx=10, pady=10, sticky='e')
entry_output_filename_tab2 = tk.Entry(tab2, width=50)
entry_output_filename_tab2.grid(row=7, column=1, padx=10, pady=10)

button_start_tab2 = tk.Button(tab2, text="작업 시작", command=lambda: start_processing2(entry_file_path_tab2, entry_sheet_name_tab2, entry_url_column_name_tab2, entry_cafe_name_column_tab2, entry_member_count_column_tab2, entry_keywords_tab2, entry_output_filename_tab2))
button_start_tab2.grid(row=8, column=1, padx=10, pady=10)

tab_control.pack(expand=1, fill="both")
root.mainloop()
