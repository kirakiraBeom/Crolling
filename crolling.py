import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import time
import pandas as pd
import os
import re  # 파일 이름에서 불필요한 문자 제거를 위한 모듈
from selenium.webdriver.common.alert import Alert

# 카테고리와 URL 매핑 (일간/주간 구분)
CATEGORY_URLS = {
    "전체": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000055&categoryChildCategoryId=&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000055&categoryChildCategoryId=&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "DIY 용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000933&categoryChildCategoryId=50000933&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000933&categoryChildCategoryId=50000933&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "램프": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000934&categoryChildCategoryId=50000934&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000934&categoryChildCategoryId=50000934&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "배터리용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000936&categoryChildCategoryId=50000936&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000936&categoryChildCategoryId=50000936&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "공기청정용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000937&categoryChildCategoryId=50000937&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000937&categoryChildCategoryId=50000937&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "세차용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000938&categoryChildCategoryId=50000938&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000938&categoryChildCategoryId=50000938&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "키용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000939&categoryChildCategoryId=50000939&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D",
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000939&categoryChildCategoryId=50000939&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "편의용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000940&categoryChildCategoryId=50000940&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D",
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000940&categoryChildCategoryId=50000940&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "오일관리": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000941&categoryChildCategoryId=50000941&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D",
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000941&categoryChildCategoryId=50000941&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "익스테리어용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000942&categoryChildCategoryId=50000942&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D",
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000942&categoryChildCategoryId=50000942&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "인테리어용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000943&categoryChildCategoryId=50000943&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D",
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000943&categoryChildCategoryId=50000943&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "전기용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000944&categoryChildCategoryId=50000944&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000944&categoryChildCategoryId=50000944&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "수납용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000945&categoryChildCategoryId=50000945&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000945&categoryChildCategoryId=50000945&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "휴대폰용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000946&categoryChildCategoryId=50000946&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000946&categoryChildCategoryId=50000946&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "타이어/휠": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000947&categoryChildCategoryId=50000947&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000947&categoryChildCategoryId=50000947&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    },
    "튜닝용품": {
        "일간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000948&categoryChildCategoryId=50000948&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P1D", 
        "주간": "https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000948&categoryChildCategoryId=50000948&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&chartRank=1&period=P7D"
    }
}

def start_crawling():
    start_rank = start_rank_entry.get()
    end_rank = end_rank_entry.get()
    category = category_combobox.get()
    period_daily = period_daily_var.get()
    period_weekly = period_weekly_var.get()

    # 카테고리 선택 확인
    if category == "------------":
        messagebox.showerror("Error", "카테고리를 선택해주세요.")
        return
    
    # 순위 입력 확인
    if not start_rank.isdigit() or not end_rank.isdigit():
        messagebox.showerror("Error", "순위 입력이 잘못되었습니다. 숫자로 입력해주세요.")
        return

    start_rank = int(start_rank)
    end_rank = int(end_rank)

    # 순위 범위 확인
    if start_rank < 1 or end_rank > 100 or start_rank > end_rank:
        messagebox.showerror("Error", "순위 범위가 잘못되었습니다. 1부터 100 사이의 숫자로 입력해주세요.")
        return

    # 일간/주간 체크박스 확인 후 크롤링 스레드 시작
    if period_daily:
        url = CATEGORY_URLS[category]["일간"]
        if url:
            threading.Thread(target=crawl_naver_shopping, args=(start_rank, end_rank, url, "일간", category)).start()
    
    if period_weekly:
        url = CATEGORY_URLS[category]["주간"]
        if url:
            threading.Thread(target=crawl_naver_shopping, args=(start_rank, end_rank, url, "주간", category)).start()

def crawl_naver_shopping(start_rank, end_rank, url, period, category):
    service = Service(r'C:\Program Files\SeleniumBasic\chromedriver.exe')
    driver = webdriver.Chrome(service=service)

    try: 
        driver.get(url)
        scroll_to_bottom(driver)

        # 정확한 날짜 정보만 추출
        date_element = driver.find_element(By.CLASS_NAME, 'periodFilter_standard__UiHDw')
        link_info_elements = date_element.find_elements(By.CLASS_NAME, 'periodFilter_link_info__lZkOv')
        for elem in link_info_elements:
            driver.execute_script("""
                var element = arguments[0];
                element.parentNode.removeChild(element);
                """, elem)

        date_text = date_element.text.strip()  # 텍스트만 추출하고 공백 제거
        date_text = re.sub(r'[<>:"/\\|?*]', '', date_text)  # 파일 이름에 사용할 수 없는 문자 제거

        data = []
        products = driver.find_elements(By.CLASS_NAME, 'imageProduct_item__KZB_F')[start_rank-1:end_rank]

        for index, product in enumerate(products):
            try:
                product_name = product.find_element(By.CLASS_NAME, 'imageProduct_title__Wdeb1').text
                store_buttons = product.find_elements(By.CLASS_NAME, 'imageProduct_btn_store__bL4eB.linkAnchor')

                if store_buttons:
                    product.click()
                    time.sleep(5)
                    driver.switch_to.window(driver.window_handles[-1])

                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.CLASS_NAME, 'buyBox_mall__Jqklw.linkAnchor._nlog_click._nlog_impression_element'))
                        )
                        mall_button = driver.find_elements(By.CLASS_NAME, 'buyBox_mall__Jqklw.linkAnchor._nlog_click._nlog_impression_element')
                        if mall_button:
                            brand_name = mall_button[0].text

                            # "네이버플러스멤버십" 제거하고 앞의 내용만 남기기
                            if "네이버플러스멤버십" in brand_name:
                                brand_name = brand_name.split("네이버플러스멤버십")[0].strip()

                            mall_button[0].click()
                            time.sleep(5)
                            driver.switch_to.window(driver.window_handles[-1])
                            time.sleep(5)

                            current_url = driver.current_url

                            # 특정 URL 확인 및 스킵
                            if any(x in current_url for x in ['11st', 'coupang', 'lotteon', 'gmarket', 'auction','window-products/play/']):
                                driver.close()
                                driver.switch_to.window(driver.window_handles[-1])
                                raise Exception("Skipped non-target URL")

                            brand_check = 'O' if 'brand.naver' in current_url else 'X'
                            data.append({'순위': start_rank + index, '제품명': product_name, '브랜드명': brand_name, 'Brand 여부': brand_check})

                            driver.close()
                            time.sleep(2)
                            driver.switch_to.window(driver.window_handles[-1])
                        else:
                            raise Exception("Mall button not found, trying alternate method")

                    except Exception:
                        # URL이 스킵되거나 mall_button이 없을 때 대체 방법 사용
                        try:
                            floating_tab = driver.find_element(By.CLASS_NAME, 'floatingTab_on__2FzR0')
                            alternate_button = floating_tab.find_elements(By.CLASS_NAME, '_nlog_click._nlog_impression_element')
                            if alternate_button:
                                alternate_button[0].click()
                                time.sleep(2)
                                driver.switch_to.window(driver.window_handles[-1])
                                time.sleep(2)

                                productList_buttons = driver.find_elements(By.CLASS_NAME, 'productList_mall_link__TrYxC.linkAnchor._nlog_click._nlog_impression_element')
                                button_clicked = False
                                
                                for productList_button in productList_buttons:
                                    try:
                                        # 브랜드명 추출
                                        brand_name = productList_button.find_element(By.TAG_NAME, 'span').text

                                        productList_button.click()
                                        time.sleep(5)
                                        driver.switch_to.window(driver.window_handles[-1])
                                        time.sleep(5)

                                        current_url = driver.current_url
                                        
                                        if any(x in current_url for x in ['11st', 'coupang', 'lotteon', 'gmarket', 'auction','window-products/play/']):
                                            driver.close()
                                            time.sleep(2)
                                            driver.switch_to.window(driver.window_handles[-1])
                                            continue  # Skip this button and try the next
                                        else:
                                            brand_check = 'O' if 'brand.naver' in current_url else 'X'
                                            data.append({'순위': start_rank + index, '제품명': product_name, '브랜드명': brand_name, 'Brand 여부': brand_check})
                                            button_clicked = True
                                            break  # Exit the loop once a valid button is found

                                    except Exception as e:
                                        print(f"Error processing product {start_rank + index}: {e}")
                                        driver.switch_to.window(driver.window_handles[-1])
                                        continue

                                if not button_clicked:
                                    data.append({'순위': start_rank + index, '제품명': product_name, '브랜드명': 'N/A', 'Brand 여부': 'X'})
                            
                            else:
                                data.append({'순위': start_rank + index, '제품명': product_name, '브랜드명': 'N/A', 'Brand 여부': 'X'})
                        except Exception as e:
                            print(f"Alternate method failed for product {start_rank + index}: {e}")
                            data.append({'순위': start_rank + index, '제품명': product_name, '브랜드명': 'N/A', 'Brand 여부': 'X'})

                    driver.close()
                    time.sleep(2)
                    driver.switch_to.window(driver.window_handles[0])
                else:
                    try:
                        brand_name = product.find_element(By.CLASS_NAME, 'imageProduct_mall__tJkQR').text
                        
                        # "네이버플러스멤버십" 제거하고 앞의 내용만 남기기
                        if "네이버플러스멤버십" in brand_name:
                            brand_name = brand_name.split("네이버플러스멤버십")[0].strip()

                        product.click()
                        time.sleep(5)
                        driver.switch_to.window(driver.window_handles[-1])
                        time.sleep(5)

                        current_url = driver.current_url
                        brand_check = 'O' if 'brand.naver' in current_url else 'X'
                        data.append({'순위': start_rank + index, '제품명': product_name, '브랜드명': brand_name, 'Brand 여부': brand_check})

                        driver.close()
                        time.sleep(2)
                        driver.switch_to.window(driver.window_handles[0])
                    except Exception as e:
                        print(f"Error processing product {start_rank + index}: {e}")
                        driver.close()
                        time.sleep(2)
                        driver.switch_to.window(driver.window_handles[0])

            except Exception as e:
                print(f"Error processing product {start_rank + index}: {e}")
                driver.switch_to.window(driver.window_handles[0])
                continue

        # 엑셀 파일 이름 생성 (중복 처리)
        output_file = f'{date_text}_{period}_{category}.xlsx'
        output_file = get_unique_filename(output_file)
        driver.quit()
        df = pd.DataFrame(data)
        df.to_excel(output_file, index=False)
        messagebox.showinfo("Success", f"엑셀 파일이 저장되었습니다: {output_file}")

    finally:
        driver.quit()


def get_unique_filename(file_path):
    """
    파일 경로가 중복될 경우, 숫자를 추가하여 고유한 파일 이름을 생성합니다.
    """
    base, extension = os.path.splitext(file_path)
    counter = 1

    # 파일이 이미 존재할 경우, 숫자를 추가하여 고유한 파일 이름 생성
    while os.path.exists(file_path):
        file_path = f"{base}({counter}){extension}"
        counter += 1
    
    return file_path

def scroll_to_bottom(driver):
    last_height = driver.execute_script("return document.body.scrollHeight")
    
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height     
        
# GUI 창을 화면 중앙에 위치시키는 함수
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = int((screen_width / 2) - (width / 2))
    y = int((screen_height / 2) - (height / 2))

    window.geometry(f"{width}x{height}+{x}+{y}")

# 메인 윈도우 생성
root = tk.Tk()
root.title("brand.naver 확인")
root.geometry("280x200")
root.attributes('-topmost', True)

# 화면 중앙에 위치시키기
center_window(root, 280, 200)
# 창 크기 고정
root.resizable(False, False)

# 시작 순위 라벨 및 입력
tk.Label(root, text="시작 순위(1~)").grid(row=0, column=0, padx=5, pady=5)
start_rank_entry = tk.Entry(root)
start_rank_entry.grid(row=0, column=1, padx=5, pady=5)
start_rank_entry.insert(0, "1")

# 끝 순위 라벨 및 입력
tk.Label(root, text="끝 순위(~100)").grid(row=1, column=0, padx=5, pady=5)
end_rank_entry = tk.Entry(root)
end_rank_entry.grid(row=1, column=1, padx=5, pady=5)
end_rank_entry.insert(0, "100")

# 카테고리 라벨 및 콤보박스
tk.Label(root, text="카테고리").grid(row=2, column=0, padx=5, pady=5)
category_combobox = ttk.Combobox(root, state="readonly")  # Disable typing
category_combobox['values'] = ["------------", "전체", "DIY 용품", "램프", "배터리용품", "공기청정용품", "세차용품", "키용품", "편의용품", "오일관리", "익스테리어용품", "인테리어용품", "전기용품", "수납용품", "휴대폰용품", "타이어/휠", "튜닝용품"]
category_combobox.grid(row=2, column=1, padx=5, pady=5)
category_combobox.current(0)

# 기간 선택 라벨 및 체크박스
tk.Label(root, text="기간 선택").grid(row=3, column=0, padx=5, pady=5)
period_daily_var = tk.BooleanVar(value=True)
period_weekly_var = tk.BooleanVar(value=False)
tk.Checkbutton(root, text="일간", variable=period_daily_var).grid(row=3, column=1, padx=0, pady=5, sticky='w')
tk.Checkbutton(root, text="주간", variable=period_weekly_var).grid(row=3, column=1, padx=60, pady=5, sticky='w')

# START 버튼을 가운데에 위치시키기 위해 columnspan 사용
start_button = tk.Button(root, text="START", command=start_crawling)
start_button.grid(row=4, column=0, columnspan=2, pady=10)

def on_window_focus(event):
    root.attributes('-topmost', False)

root.bind("<FocusOut>", on_window_focus)

# GUI 실행
root.mainloop()
