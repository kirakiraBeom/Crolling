import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from threading import Thread
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException

# Chrome WebDriver 설정
driver_path = "C:\\Program Files\\SeleniumBasic\\chromedriver.exe"
service = Service(driver_path)

# 스마트스토어와 브랜드스토어 URL 매핑
smart_urls = {
    "유투카": "https://smartstore.naver.com/youtocar",
    "세이보링": "https://smartstore.naver.com/savoring", 
    "안녕하십니카": "https://smartstore.naver.com/annyeongcar",
    "킨톤": "https://smartstore.naver.com/create1988",
    "이지스오토랩": "https://smartstore.naver.com/aegisautolab",
    "기가차스토어": "https://smartstore.naver.com/gigachastore",
    "테스트드라이브(카마루)": "https://smartstore.naver.com/testdrive",
    "트니르": "https://smartstore.naver.com/tenir_",
    "더레이즈": "https://smartstore.naver.com/theraise",
    "발상코퍼레이션": "https://smartstore.naver.com/superliving",
    "데코M(엠노블)": "https://smartstore.naver.com/deco-m",
}

brand_urls = {
    "메이튼": "https://brand.naver.com/mayton",
    "벤딕트": "https://brand.naver.com/vendict",
    "케이엠모터스": "https://brand.naver.com/kmmotors",
    "지엠지모터스(꾸미자닷컴)": "https://brand.naver.com/gmzmotors",
    "본투로드": "https://brand.naver.com/motorlife",
    "가온": "https://brand.naver.com/gaoncarmat",
    "아임반": "https://brand.naver.com/aimban",
    "돌비웨이": "https://brand.naver.com/dolbiway",
    "카멜레온360": "https://brand.naver.com/chameleon360",
    "주파집": "https://brand.naver.com/jupazip",
}

# 브랜드와 URL 병합
all_urls = {**smart_urls, **brand_urls}

def open_browser():
    """크롬 브라우저를 열고 빈 탭 생성"""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")  # 브라우저 창을 최대화하여 시작
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("about:blank")  # 빈 탭 열기
    return driver

def open_new_tab(driver, url):
    """새 탭을 열고 지정된 URL로 이동"""
    driver.execute_script("window.open('');")  # 새 탭 열기
    driver.switch_to.window(driver.window_handles[-1])  # 새로 열린 탭으로 이동
    driver.get(url)

def get_element_text_by_class(driver, parent_class, child_class, index=0):
    """지정된 부모 클래스 안의 자식 클래스 요소에서 텍스트 가져오기"""
    try:
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, parent_class))
        )
        element_text = parent_element.find_elements(By.CLASS_NAME, child_class)[index].text
        return element_text
    except Exception as e:
        return None

def get_store_grade(driver, brand):
    """스토어 등급 가져오기"""
    if brand in brand_urls:
        # 브랜드스토어의 경우 등급 없음
        return None
    else:
        return get_element_text_by_class(driver, "_2woROoaPZC", "_3CfLtIh1fI")

def get_customer_interest(driver, brand):
    """관심 고객 수 가져오기 (숫자와 쉼표만)"""
    if brand in brand_urls:
        full_text = get_element_text_by_class(driver, "_3aNsjos9K5", "_3e458DWUPL")
    else:
        full_text = get_element_text_by_class(driver, "_14Ezl7R3c-", "_3KDc7jvaa-")
        
    if full_text:
        # 숫자와 쉼표만 추출
        numbers_only = re.findall(r"[\d,]+", full_text)
        if numbers_only:
            return numbers_only[0]
    return None

def get_list_count(driver, brand):
    """목록 수 가져오기"""
    try:
        if brand in brand_urls:
            parent_class = "gpz-TGlihm"
            child_class = "_1xbiPyV_cm"
            sub_class = "_3dg1lTjfos"
        else:
            parent_class = "XgIBU904A2"
            child_class = "ZjAm7j87Us"
            sub_class = "_1nzrcyiki9"
            
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, parent_class))
        )
        child_element = parent_element.find_element(By.CLASS_NAME, child_class)
        list_items = child_element.find_elements(By.CLASS_NAME, sub_class)
        return len(list_items)
    except Exception as e:
        return None

def click_more_button(driver):
    """추가 정보 로딩을 위한 버튼 클릭 (없으면 패스)"""
    try:
        more_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "_3ryPAjhmjZ"))
        )
        more_button.click()
        time.sleep(3)  # 클릭 후 페이지가 로드되도록 대기
        return True
    except Exception as e:
        print("더보기 버튼을 찾을 수 없음. 다음 단계로 넘어갑니다.")
        return True  # 더보기 버튼이 없을 경우에도 다음 단계로 진행

def get_category_list(driver):
    """카테고리 목록 가져오기"""
    try:
        category_list = []
        # _1J2oAxZvAG 클래스 내부의 _3AV7RVieRB 클래스 찾기
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "_1J2oAxZvAG"))
        )
        child_elements = parent_element.find_elements(By.CLASS_NAME, "_3AV7RVieRB")
        
        for element in child_elements:
            # 각 _3AV7RVieRB 클래스 안의 여러 _2jm5JW3D5W클래스의 <a> 태그에서 텍스트 추출
            sub_elements = element.find_elements(By.CLASS_NAME, "_2jm5JW3D5W")
            for sub_element in sub_elements:
                category_text = sub_element.find_element(By.TAG_NAME, "a").text
                # 텍스트에서 불필요한 부분을 제거
                clean_text = category_text.replace("하위 메뉴 있음", "").replace("전체상품", "").replace("전체 상품", "").strip()
                if clean_text:  # 빈 문자열이 아닌 경우만 추가
                    category_list.append(clean_text)
        
        return category_list
    except Exception as e:
        return None

def go_to_next_page(driver):
    """카테고리 목록에서 '전체상품' 또는 '전체 상품' 링크로 이동하여 인기도순 정렬 클릭"""
    try:
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "_1J2oAxZvAG"))
        )
        # _3AV7RVieRB 클래스 안의 _2jm5JW3D5W 클래스에서 "전체상품" 또는 "전체 상품" 링크 찾기
        category_elements = parent_element.find_elements(By.CLASS_NAME, "_3AV7RVieRB")
        for element in category_elements:
            sub_elements = element.find_elements(By.CLASS_NAME, "_2jm5JW3D5W")
            for sub_element in sub_elements:
                category_text = sub_element.find_element(By.TAG_NAME, "a").text
                # '전체상품' 또는 '전체 상품'이 포함된 텍스트를 확인
                if "전체상품" in category_text or "전체 상품" in category_text:
                    category_url = sub_element.find_element(By.TAG_NAME, "a").get_attribute("href")
                    driver.get(category_url)  # '전체상품' URL로 이동
                    time.sleep(3)  # 페이지 로딩 대기

                    # "인기도순" 텍스트를 가진 버튼을 찾고 클릭
                    popularity_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='인기도순']"))
                    )
                    popularity_button.click()
                    time.sleep(3)  # 클릭 후 대기
                    return True
    except Exception as e:
        print(f"오류 발생: {e}")
        return False

def get_top_10_products(driver, brand):
    """인기 Top 10 제품명 가져오기"""
    try:
        if brand in brand_urls:
            parent_class = "_1bijXwxRam"
            child_class = "_3iW9G4pEbm"
            product_class = "_3pA0Duwrhw"
        else:
            parent_class = "wOWfwtMC_3"
            child_class = "flu7YgFW2k"
            product_class = "_26YxgX-Nu5"
        
        parent_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, parent_class))
        )
        top_10_products = []
        
        for parent_element in parent_elements:
            child_elements = parent_element.find_elements(By.CLASS_NAME, child_class)
            for child_element in child_elements:
                if len(top_10_products) >= 10:
                    break
                try:
                    product_name = child_element.find_element(By.CLASS_NAME, product_class).text
                    top_10_products.append(product_name)
                except:
                    continue
            if len(top_10_products) >= 10:
                break
        
        return top_10_products
    except Exception as e:
        return None

def click_latest_button(driver):
    """'최신등록순' 버튼 클릭"""
    try:
        latest_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[text()='최신등록순']"))
        )
        latest_button.click()
        time.sleep(5)  # 클릭 후 페이지 로드 대기
        return True
    except Exception as e:
        return False

def get_latest_products(driver, brand):
    """최근 등록된 제품명 가져오기 (특정 조건을 만족하는 제품만)"""
    try:
        if brand in brand_urls:
            parent_class = "_1bijXwxRam"
            child_class = "ZdiAiTrQWZ"
            sub_class = "_3iW9G4pEbm"
            product_class = "_3pA0Duwrhw"
            condition_class = "_3SoCq43A_l"
        else:
            parent_class = "wOWfwtMC_3"
            child_class = "flu7YgFW2k"
            sub_class = "_2Ngqhufxxc"
            product_class = "_26YxgX-Nu5"
            condition_class = "_2Ngqhufxxc"
        
        parent_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, parent_class))
        )
        latest_products = []
        
        for parent_element in parent_elements:
            child_elements = parent_element.find_elements(By.CLASS_NAME, child_class)
            for child_element in child_elements:
                try:
                    # 조건 클래스가 있는지 확인
                    if child_element.find_elements(By.CLASS_NAME, condition_class):
                        product_name = child_element.find_element(By.CLASS_NAME, product_class).text
                        latest_products.append(product_name)
                except:
                    continue
        
        return latest_products
    except Exception as e:
        return None

def get_total_product_count(driver):
    """전체 상품 갯수 가져오기"""
    try:
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "_3VrqrkLvIc"))
        )
        # 두 번째 _3-RHMpXayH 클래스 요소를 찾음
        sub_elements = parent_element.find_elements(By.CLASS_NAME, "_3-RHMpXayH")
        if len(sub_elements) >= 2:
            count_element = sub_elements[1].find_element(By.CLASS_NAME, "_6lgM26zUO6")
            product_count = count_element.find_element(By.TAG_NAME, "strong").text
            return product_count
        else:
            return None
    except Exception as e:
        return None

def close_browser(driver):
    """브라우저 닫기"""
    driver.quit()

def scrape_brand_data(driver, brand_name, url):
    """지정된 브랜드 URL에서 데이터를 수집하여 딕셔너리로 반환"""
    open_new_tab(driver, url)  # 새로운 탭에서 브랜드 URL 열기
    data = {
        "스토어 등급": [],
        "관심 고객 수": [],
        "목록 수": [],
        "카테고리 목록": [],
        "인기 Top 10 제품명": [],
        "최신 등록된 제품명": [],
        "전체 상품 갯수": []
    }

    # 스토어 등급 가져오기
    store_grade = get_store_grade(driver, brand_name)
    if store_grade:
        data["스토어 등급"].append(store_grade)
    
    # 관심 고객 수 가져오기
    customer_interest = get_customer_interest(driver, brand_name)
    if customer_interest:
        data["관심 고객 수"].append(customer_interest)
    
    # 목록 수 가져오기
    list_count = get_list_count(driver, brand_name)
    if list_count is not None:
        data["목록 수"].append(list_count)
    
    # 더보기 버튼 클릭 (버튼 없으면 다음으로 넘어감)
    click_more_button(driver)
    
    # 카테고리 목록 가져오기
    category_list = get_category_list(driver)
    if category_list:
        data["카테고리 목록"] = category_list
    
    # 다음 페이지로 이동 및 인기도순 정렬 클릭
    if go_to_next_page(driver):
        # 인기 Top 10 제품명 가져오기
        top_10_products = get_top_10_products(driver, brand_name)
        if top_10_products:
            data["인기 Top 10 제품명"] = top_10_products

        # 최신등록순 클릭 및 최신 제품 가져오기
        if click_latest_button(driver):
            latest_products = get_latest_products(driver, brand_name)
            if latest_products:
                data["최신 등록된 제품명"] = latest_products

        # 전체 상품 갯수 가져오기
        total_count = get_total_product_count(driver)
        if total_count:
            data["전체 상품 갯수"].append(total_count)
    
    driver.close()  # 현재 탭 닫기
    driver.switch_to.window(driver.window_handles[0])  # 첫 번째 탭으로 전환
    return data

def save_to_excel(all_data):
    """수집된 데이터를 엑셀 파일로 저장"""
    file_name = "경쟁사 네이버스토어_data.xlsx"
    try:
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for brand_name, data in all_data.items():
                df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in data.items()]))
                df.to_excel(writer, sheet_name=brand_name, index=False)
        messagebox.showinfo("저장 완료", f"{file_name} 저장 완료.")
    except (PermissionError, InvalidFileException):
        messagebox.showerror("저장 실패", f"엑셀 파일이 실행 중이거나 접근할 수 없습니다. '{file_name}'을 닫고 다시 시도하세요.")

def start_scraping(selected_brands):
    """선택된 브랜드를 기준으로 스크래핑 시작"""
    driver = open_browser()
    all_data = {}

    for brand in selected_brands:
        data = scrape_brand_data(driver, brand, all_urls[brand])
        all_data[brand] = data

    save_to_excel(all_data)
    close_browser(driver)

def run_scraping_in_thread(selected_brands):
    """별도의 스레드에서 스크래핑 작업 수행"""
    scraping_thread = Thread(target=start_scraping, args=(selected_brands,))
    scraping_thread.start()

def create_gui():
    """Tkinter GUI를 생성하여 브랜드 선택 UI 제공"""
    root = tk.Tk()
    root.title("브랜드 선택")
    root.geometry("500x420")  # 창 크기 설정
    root.resizable(False, False)  # 창 크기 변경 불가

    selected_brands = []
    
    def on_select():
        selected_brands.clear()
        for i, var in enumerate(check_vars):
            if var.get():
                selected_brands.append(brand_list[i])
        if selected_brands:
            run_scraping_in_thread(selected_brands)

    def toggle_all():
        """전체 선택/해제 버튼 동작"""
        new_state = not all(var.get() for var in check_vars)
        for var in check_vars:
            var.set(new_state)
        toggle_all_button.config(text="전체 해제" if new_state else "전체 선택")

    def toggle_smartstore():
        """스마트스토어 전체 선택/해제"""
        new_state = not all(check_vars[i].get() for i in range(len(smart_urls)))
        for i in range(len(smart_urls)):
            check_vars[i].set(new_state)
        smartstore_button.config(text="전체 해제" if new_state else "전체 선택")

    def toggle_brandstore():
        """브랜드스토어 전체 선택/해제"""
        new_state = not all(check_vars[i].get() for i in range(len(smart_urls), len(check_vars)))
        for i in range(len(smart_urls), len(check_vars)):
            check_vars[i].set(new_state)
        brandstore_button.config(text="전체 해제" if new_state else "전체 선택")

    # 브랜드 목록 생성
    brand_list = list(all_urls.keys())
    check_vars = [tk.BooleanVar() for _ in brand_list]

    # 스타일 설정
    style = ttk.Style()
    style.configure("TButton", font=("Arial", 10, "bold"), foreground="white", background="#4CAF50")
    style.configure("TCheckbutton", font=("Arial", 10))
    style.configure("TLabel", font=("Arial", 12, "bold"), background="#3B5998", foreground="white")
    style.configure("TFrame", background="#f0f0f0")

    # 프레임 생성
    top_frame = ttk.Frame(root, padding="10 10 10 10", relief="solid", borderwidth=1)
    top_frame.pack(pady=5, fill='both', expand=True)
    
    left_frame = ttk.Frame(top_frame, relief="solid", borderwidth=1, padding="5 5 5 5")
    right_frame = ttk.Frame(top_frame, relief="solid", borderwidth=1, padding="5 5 5 5")
    left_frame.pack(side='left', padx=20, pady=10, fill='y')
    right_frame.pack(side='right', padx=20, pady=10, fill='y')

    # 스마트스토어 제목
    ttk.Label(left_frame, text="스마트스토어").grid(row=0, column=0, sticky='w')
    # 스마트스토어 체크박스
    for i, brand in enumerate(smart_urls.keys()):
        ttk.Checkbutton(left_frame, text=brand, variable=check_vars[i]).grid(row=i+1, column=0, sticky='w')

    # 브랜드스토어 제목
    ttk.Label(right_frame, text="브랜드스토어").grid(row=0, column=0, sticky='w')
    # 브랜드스토어 체크박스
    for j, brand in enumerate(brand_urls.keys(), start=len(smart_urls)):
        ttk.Checkbutton(right_frame, text=brand, variable=check_vars[j]).grid(row=j-len(smart_urls)+1, column=0, sticky='w')

    # 아래의 버튼 프레임
    bottom_frame = ttk.Frame(root)
    bottom_frame.pack(pady=10)

    # 스마트스토어 전체 선택/해제 버튼
    smartstore_button = ttk.Button(left_frame, text="전체 선택", command=toggle_smartstore)
    smartstore_button.grid(row=len(smart_urls) + 1, column=0, pady=10)

    # 브랜드스토어 전체 선택/해제 버튼
    brandstore_button = ttk.Button(right_frame, text="전체 선택", command=toggle_brandstore)
    brandstore_button.grid(row=len(brand_urls) + 1, column=0, pady=10)

    # 전체 선택/해제 버튼
    toggle_all_button = ttk.Button(bottom_frame, text="전체 선택", command=toggle_all)
    toggle_all_button.grid(row=0, column=0, padx=5)

    start_button = ttk.Button(bottom_frame, text="시작", command=on_select)
    start_button.grid(row=0, column=1, padx=5)

    root.mainloop()

# GUI 시작
create_gui()
