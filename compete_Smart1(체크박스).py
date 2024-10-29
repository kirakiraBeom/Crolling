import json
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import openpyxl

# JSON 파일 읽기
with open('selectors.json', 'r', encoding='utf-8') as file:
    selectors = json.load(file)

# 결과를 저장할 리스트
results = []
driver = None  # 전역 driver 선언

def crawl_homepage(homepage_name, url, a_selector, b_selector, group):
    global driver  # 전역 driver 사용
    try:
        driver.get(url)
        print(f"{homepage_name} 페이지 로딩 중...")

        # A 그룹의 경우
        if group == 'A':
            # 메인 대문 갯수
            buttons = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@class='_3Te07yM0Z_']//div[@class='ZjAm7j87Us']/a/span"))
            )
            button_count = len(buttons)

            # 관심 고객 수 가져오기
            interested_customers = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[@class='_3KDc7jvaa-']")
                )
            ).text  # 전체 텍스트 가져오기

            # 텍스트에서 숫자 추출
            interested_customers_count = "".join(filter(str.isdigit, interested_customers))  # 숫자만 추출
            if not interested_customers_count:  # 숫자가 없으면
                interested_customers_count = "N/A"  # 값이 없을 경우

            # 전체상품 URL 가져오기
            try:
                parent_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "_1J2oAxZvAG"))
                )
                category_elements = parent_element.find_elements(By.CLASS_NAME, "_3AV7RVieRB")
                for element in category_elements:
                    sub_elements = element.find_elements(By.CLASS_NAME, "_2jm5JW3D5W")
                    for sub_element in sub_elements:
                        category_text = sub_element.find_element(By.TAG_NAME, "a").text
                        # '전체상품' 또는 '전체 상품'이 포함된 텍스트를 확인
                        if "전체상품" in category_text or "전체 상품" in category_text:
                            category_url = sub_element.find_element(By.TAG_NAME, "a").get_attribute("href")
                            driver.get(category_url)  # '전체상품' URL로 이동
                            break  # URL로 이동 후 루프 종료
            except Exception as e:
                print(f"전체상품 URL 찾기 중 오류 발생: {e}")

            # 전체 상품 페이지에서 버튼 클릭
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//li[@class='_1GSO93arMl']//button"))
            ).click()

            # 크롤링할 리스트 항목 가져오기
            product_items = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//ul[contains(@class, 'wOWfwtMC_3 _3cLKMqI7mI')]/li"))
            )[:10]  # 첫 번째부터 열 번째까지

            # 제품명 추출
            product_names = [item.find_element(By.CLASS_NAME, "_26YxgX-Nu5").text for item in product_items]
            print(f"A 그룹 제품명: {product_names}")

        # B 그룹의 경우
        elif group == 'B':
            # 메인 대문 갯수
            buttons = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[@class='_1xbiPyV_cm']//ul/li"))
            )
            button_count = len(buttons)

            # 관심 고객 수 가져오기
            interested_customers = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[@class='_3e458DWUPL']")
                )
            ).text  # 전체 텍스트 가져오기

            # 텍스트에서 숫자 추출
            interested_customers_count = "".join(filter(str.isdigit, interested_customers))  # 숫자만 추출
            if not interested_customers_count:  # 숫자가 없으면
                interested_customers_count = "N/A"  # 값이 없을 경우

            # 전체상품 URL 가져오기
            try:
                parent_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "_1J2oAxZvAG"))
                )
                category_elements = parent_element.find_elements(By.CLASS_NAME, "_3AV7RVieRB")
                for element in category_elements:
                    sub_elements = element.find_elements(By.CLASS_NAME, "_2jm5JW3D5W")
                    for sub_element in sub_elements:
                        category_text = sub_element.find_element(By.TAG_NAME, "a").text
                        # '전체상품' 또는 '전체 상품'이 포함된 텍스트를 확인
                        if "전체상품" in category_text or "전체 상품" in category_text:
                            category_url = sub_element.find_element(By.TAG_NAME, "a").get_attribute("href")
                            driver.get(category_url)  # '전체상품' URL로 이동
                            break  # URL로 이동 후 루프 종료
            except Exception as e:
                print(f"전체상품 URL 찾기 중 오류 발생: {e}")

            # 전체 상품 페이지에서 버튼 클릭
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//li[@class='_1GSO93arMl']//button"))
            ).click()

            # 크롤링할 리스트 항목 가져오기
            product_items = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//ul[contains(@class, '_1bijXwxRam _26gVNSMpdB')]/li"))
            )[:10]  # 첫 번째부터 열 번째까지

            # 제품명 추출
            product_names = [item.find_element(By.CLASS_NAME, "_3pA0Duwrhw").text for item in product_items]
            print(f"B 그룹 제품명: {product_names}")

        # 결과 추가
        results.append({
            "Homepage": homepage_name,
            "Button Count": button_count,
            "Store Grade": "",
            "Interested Customers": interested_customers_count,
            "Top 10 Products": ", ".join(product_names),  # 인기 제품명 추가
            "Latest Products": "",
            "Category List": "",
            "Total Products": ""
        })

    except NoSuchElementException as e:
        print(f"{homepage_name}에서 요소를 찾을 수 없습니다. 오류: {e}")
        results.append({"Homepage": homepage_name, "Button Count": 0, "Store Grade": "", "Interested Customers": "N/A",
                       "Top 10 Products": "", "Latest Products": "", "Category List": "", "Total Products": ""})
    except TimeoutException as e:
        print(f"{homepage_name} 로딩 중 시간 초과. 오류: {e}")
        results.append({"Homepage": homepage_name, "Button Count": 0, "Store Grade": "", "Interested Customers": "N/A",
                       "Top 10 Products": "", "Latest Products": "", "Category List": "", "Total Products": ""})
    except Exception as e:
        print(f"{homepage_name} 크롤링 중 오류 발생: {e}")
        results.append({"Homepage": homepage_name, "Button Count": 0, "Store Grade": "", "Interested Customers": "N/A",
                       "Top 10 Products": "", "Latest Products": "", "Category List": "", "Total Products": ""})


def save_to_excel(data):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Crawling Results"

    # 헤더 추가
    sheet.append(["경쟁사", "메인대문갯수", "스토어등급", "관심고객수",
                 "인기top10 제품명", "최신등록순 제품명", "카테고리목록", "전체상품수"])

    # 데이터 추가
    for result in data:
        sheet.append([result["Homepage"], result["Button Count"], result["Store Grade"], result["Interested Customers"],
                     result["Top 10 Products"], result["Latest Products"], result["Category List"], result["Total Products"]])

    workbook.save("crawling_results.xlsx")
    print("크롤링 결과가 'crawling_results.xlsx' 파일에 저장되었습니다.")

def start_crawling():
    global results, driver  # 결과와 driver를 전역 변수로 사용
    results = []  # 기존 결과 초기화

    # 크롬 드라이버 초기화
    options = Options()
    options.add_argument("--start-maximized")  # 최대화된 상태로 시작
    options.add_argument("--ignore-certificate-errors")  # SSL 인증서 오류 무시
    driver = webdriver.Chrome(service=Service(
        "C:\\Program Files\\SeleniumBasic\\chromedriver.exe"), options=options)

    selected_competitors = [competitor for var, competitor in zip(
        competitor_vars, selectors.items()) if var.get()]

    for name, competitor in selected_competitors:
        url = competitor["url"]
        a_selector = competitor.get("a_selector", None)
        b_selector = competitor.get("b_selector", None)

        # A 그룹은 a_selector 사용, B 그룹은 b_selector 사용
        if 'smartstore' in url:
            crawl_homepage(name, url, a_selector, None, 'A')
        else:
            crawl_homepage(name, url, None, b_selector, 'B')

    # 웹 드라이버 종료
    driver.quit()

    # 결과를 .xlsx 파일로 저장
    save_to_excel(results)

    # 결과 표시
    result_text = "\n".join([f"{result['Homepage']}: {result['Button Count']} 개, 관심고객수: {result['Interested Customers']}" for result in results])
    messagebox.showinfo("크롤링 결과", result_text)

# 메인 윈도우 설정
root = tk.Tk()
root.title("네이버 쇼핑 크롤러")

# 경쟁사 체크박스 프레임
competitor_frame = tk.Frame(root)
competitor_frame.pack(pady=10)

# 경쟁사 체크박스
competitor_vars = []
for index, (name, competitor) in enumerate(selectors.items()):
    var = tk.BooleanVar()
    competitor_vars.append(var)
    check = tk.Checkbutton(competitor_frame, text=name,
                           variable=var, font=("Arial", 12))
    check.grid(row=index // 3, column=index % 3, padx=10, pady=5, sticky='w')

# 전체 선택 버튼
select_all_button = tk.Button(root, text="전체 선택", command=lambda: [
                              var.set(True) for var in competitor_vars], font=("Arial", 12))
select_all_button.pack(pady=5)

# 전체 해제 버튼
deselect_all_button = tk.Button(root, text="전체 해제", command=lambda: [
                                var.set(False) for var in competitor_vars], font=("Arial", 12))
deselect_all_button.pack(pady=5)

# 크롤링 시작 버튼
start_button = tk.Button(
    root, text="크롤링 시작", command=start_crawling, font=("Arial", 12))
start_button.pack(pady=15)

# GUI 실행
root.mainloop()
