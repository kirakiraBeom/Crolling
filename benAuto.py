import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 웹드라이버 설정
driver = webdriver.Chrome()  # Chrome 드라이버 경로 설정 필요

# 네이버 쇼핑 페이지로 이동
driver.get('https://search.shopping.naver.com/best/category/click?categoryCategoryId=50000055&categoryChildCategoryId=&categoryDemo=A00&categoryMidCategoryId=50000055&categoryRootCategoryId=50000008&period=P1D&tr=nwbhi')

# 데이터 저장할 리스트
data = []

# 스크롤 및 크롤링 로직
for i in range(1, 101):
    # 상품 요소를 찾고 클릭
    product_xpath = f'//div[@class="product_list"]/div[{i}]/a'
    try:
        product = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, product_xpath)))
        product.click()
        
        # 판매처 확인
        seller_buttons = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//button[contains(text(), "판매처")]')))
        if seller_buttons:
            seller_buttons[0].click()
            
            # 가장 상단 페이지 클릭
            top_seller = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="seller_list"]/div[1]/a')))
            top_seller.click()
            
            # URL 확인
            if driver.current_url.startswith('https://brand.naver.com'):
                product_name = driver.find_element(By.XPATH, '//h1').text
                brand_name = driver.find_element(By.XPATH, '//div[@class="brand_name"]').text
                data.append([i, product_name, brand_name])
            driver.back()  # 이전 페이지로 돌아가기
        else:
            # 판매처가 한 곳인 경우 클릭
            product.click()
            
            # URL 확인
            if driver.current_url.startswith('https://brand.naver.com'):
                product_name = driver.find_element(By.XPATH, '//h1').text
                brand_name = driver.find_element(By.XPATH, '//div[@class="brand_name"]').text
                data.append([i, product_name, brand_name])
            driver.back()  # 이전 페이지로 돌아가기
    except Exception as e:
        print(f"Error at product {i}: {e}")

# 데이터프레임으로 변환 후 엑셀 저장
df = pd.DataFrame(data, columns=['순위', '제품 이름', '브랜드명'])
df.to_excel('naver_shopping_data.xlsx', index=False)

# 드라이버 종료
driver.quit()
