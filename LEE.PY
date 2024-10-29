import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By  # 업데이트된 By 모듈 사용
from bs4 import BeautifulSoup
import requests
import time  # 페이지 로딩 대기용

# 브랜드와 스마트스토어 URL 설정
brand_urls = {
    "메이튼": "https://brand.naver.com/mayton",
    "벤딕트": "https://brand.naver.com/vendict",
    "케이엠모터스": "https://brand.naver.com/kmmotors",
    # 생략...
}

smartstore_urls = {
    "유투카": "https://smartstore.naver.com/youtocar",
    "세이보링": "https://smartstore.naver.com/savoring",
    "안녕하십니카": "https://smartstore.naver.com/annyeongcar",
    # 생략...
}

# 크롬드라이버 경로 설정
driver_path = "C:\\Program Files\\SeleniumBasic\\chromedriver.exe"
service = Service(driver_path)

# 브랜드 URL 크롤링 함수
def crawl_brand(brand_name, url):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        # 페이지 구조에 따라 상품 정보 추출
        product_list = soup.find_all('div', {'class': 'product'})  # 예시 구조, 실제로 페이지 구조에 맞게 수정 필요
        for product in product_list:
            product_name = product.find('span', {'class': 'name'}).text
            price = product.find('span', {'class': 'price'}).text
            print(f'Brand: {brand_name}, Product: {product_name}, Price: {price}')
    else:
        print(f"Failed to retrieve {url}. Status code: {response.status_code}")

# 스마트스토어 URL 크롤링 함수
def crawl_smartstore(store_name, url):
    try:
        driver = webdriver.Chrome(service=service)
        driver.get(url)
        time.sleep(5)  # 페이지가 완전히 로딩될 때까지 대기

        # 상품 정보 크롤링 예시 (페이지 구조에 맞게 수정 필요)
        product_list = driver.find_elements(By.CLASS_NAME, 'product')  # 업데이트된 방식
        for product in product_list:
            product_name = product.find_element(By.CLASS_NAME, 'name').text
            price = product.find_element(By.CLASS_NAME, 'price').text
            print(f'Store: {store_name}, Product: {product_name}, Price: {price}')
    except Exception as e:
        print(f"An error occurred while crawling {store_name}: {e}")
    finally:
        driver.quit()

# GUI 함수
def start_crawling():
    selected_option = brand_var.get()
    
    if selected_option in brand_urls:
        # 브랜드 URL 크롤링
        url = brand_urls[selected_option]
        crawl_brand(selected_option, url)
    elif selected_option in smartstore_urls:
        # 스마트스토어 URL 크롤링
        url = smartstore_urls[selected_option]
        crawl_smartstore(selected_option, url)
    else:
        messagebox.showerror("Error", "유효한 브랜드나 스토어를 선택하세요.")

# GUI 설정
root = tk.Tk()
root.title("크롤링 브랜드 선택")

# 라벨
label = tk.Label(root, text="크롤링할 브랜드 또는 스마트스토어를 선택하세요:")
label.pack(pady=10)

# 브랜드 선택을 위한 라디오 버튼
brand_var = tk.StringVar()
for brand in list(brand_urls.keys()) + list(smartstore_urls.keys()):
    rb = tk.Radiobutton(root, text=brand, variable=brand_var, value=brand)
    rb.pack(anchor='w')

# 크롤링 시작 버튼
start_button = tk.Button(root, text="크롤링 시작", command=start_crawling)
start_button.pack(pady=20)

# GUI 실행
root.mainloop()
