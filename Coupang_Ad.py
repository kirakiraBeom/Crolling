import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
from functools import reduce  # reduce 함수 임포트

# Chrome WebDriver 설정
driver_path = "C:\\Program Files\\SeleniumBasic\\chromedriver.exe"
service = Service(driver_path)
driver = webdriver.Chrome(service=service)
# 창 최대화 설정
driver.maximize_window()

def open_login_page():
    """쿠팡 광고 로그인 페이지 열기"""
    driver.get("https://advertising.coupang.com/user/login")
    
def click_login_button():
    """로그인 버튼 클릭"""
    login_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//a[@href="/user/supplier/authorization"]'))
    )
    login_button.click()
    time.sleep(3)  # 페이지 로드 대기
    
def enter_username(username):
    """사용자 아이디 입력"""
    username_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.NAME, 'username'))
    )
    username_input.send_keys(username)

def wait_for_user_login():
    """사용자가 비밀번호 입력하고 로그인 버튼 누르기를 최대 3분 대기"""
    try:
        login_button = WebDriverWait(driver, 180).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@type="submit"]'))
        )
        login_button.click()
    except:
        print("로그인 시간 초과. 사용자가 로그인하지 않았습니다.")
        driver.quit()

def wait_for_ad_management_menu():
    """광고 관리 메뉴가 나타날 때까지 대기 후 클릭"""
    ad_management_menu = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//a[text()="광고관리"]'))
    )
    driver.execute_script("arguments[0].click();", ad_management_menu)
    time.sleep(3)  # 페이지 로드 대기
    
def wait_for_second_button_click():
    """사용자가 두 번째 버튼을 클릭할 때까지 대기"""
    try:
        second_button = WebDriverWait(driver, 1800).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.ant-btn.ant-btn-primary.sc-1php333-3.egZXCw'))
        )
        print("두 번째 버튼이 나타났습니다. 클릭을 기다리는 중...")

        WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ant-btn.ant-btn-primary.sc-1php333-3.egZXCw')))
        WebDriverWait(driver, 180).until_not(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.ant-btn.ant-btn-primary.sc-1php333-3.egZXCw'))
        )
        print("사용자가 두 번째 버튼을 클릭했습니다.")
        time.sleep(2)

    except Exception as e:
        print(f"버튼을 감지하거나 클릭을 기다리는 동안 문제가 발생했습니다. 에러: {e}")
        driver.quit()

def select_value_20():
    """페이지 당 20개 옵션 선택 후 3초 대기"""
    try:
        select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//select[@aria-label="rows per page"]'))
        )
        select = Select(select_element)
        select.select_by_value('20')  # value가 20인 옵션 선택
        print("옵션 20개 선택 완료.")
        time.sleep(2)

    except:
        print("옵션 선택 실패.")

def click_sorting_icon_and_wait_for_loading():
    """페이지 중간까지 스크롤한 후 정렬 아이콘 클릭 및 로딩 상태 대기"""
    try:
        # 페이지 중간까지 스크롤
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight / 2);")
        time.sleep(1)  # 스크롤 후 잠깐 대기
        
        # 요소가 클릭 가능할 때까지 대기
        sorting_icon = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.sc-dsrmhn-0.dkewmr.sorting-icon'))
        )
        
        # Actions 객체 생성 및 요소로 이동 후 클릭
        actions = ActionChains(driver)
        actions.move_to_element(sorting_icon).click().perform()

        print("정렬 아이콘 클릭 완료")
        time.sleep(10)  # 로딩 대기
        
    except Exception as e:
        print(f"정렬 아이콘 클릭 중 문제가 발생했습니다: {e}")

def get_unique_filename(filename, extension):
    """파일 이름이 중복될 경우 번호를 붙여 고유한 파일 이름 생성"""
    base_name = filename
    counter = 1
    while os.path.exists(f"{filename}.{extension}"):
        filename = f"{base_name}_{counter}"
        counter += 1
    return f"{filename}.{extension}"

def extract_data_to_excel():
    """캠페인 데이터를 추출하여 엑셀 파일로 저장"""
    try:
        time.sleep(3)
        # 데이터를 저장할 리스트
        data = []

        # rt-tbody 클래스 내부에서 rt-tr-group을 찾음
        tbody = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'rt-tbody'))
        )
        rows = tbody.find_elements(By.CLASS_NAME, 'rt-tr-group')

        for row in rows:
            # 캠페인 이름 찾기 (HTML 구조에 맞게 수정)
            campaign_name = row.find_element(By.CSS_SELECTOR, '.rt-td.rthfc-td-fixed.rthfc-td-fixed-left span').text

            # "1P"이라는 단어가 캠페인 이름에 포함되어 있는지 확인
            if "1P" in campaign_name:
                # 정확히 'rt-td' 클래스만 가지고 있는 요소를 찾음
                rt_tds = row.find_elements(By.CSS_SELECTOR, '.rt-td:not([class*=" "])')
                if len(rt_tds) >= 14:  # 최소 14개의 요소가 있는지 확인
                    ad_spend = rt_tds[4].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    conversion_sales = rt_tds[5].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    ad_profit_rate = rt_tds[6].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    impressions = rt_tds[7].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    clicks = rt_tds[8].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    ctr = rt_tds[9].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    conversion_sales_count = rt_tds[10].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    conversion_rate = rt_tds[11].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    conversion_orders = rt_tds[12].find_element(By.CLASS_NAME, 'flex-item.box--content.ar.text--flex-ellipsis').text
                    start_date = rt_tds[13].text

                    # 데이터를 리스트에 추가
                    data.append([
                        campaign_name, ad_spend, conversion_sales, ad_profit_rate, impressions, clicks, ctr,
                        conversion_sales_count, conversion_rate, conversion_orders, start_date
                    ])

        # 데이터프레임으로 변환
        df = pd.DataFrame(data, columns=[
            '캠페인 이름', '집행광고비', '광고 전환 매출', '광고수익률', '노출수', '클릭수', '클릭률',
            '광고전환판매수', '전환율', '광고 전환 주문수', '시작날짜'
        ])

        # 제거할 키워드 목록 정의
        remove_keywords = ["지난 달 :", "지난 주 :", "최근 30일 :"]

        # span 태그에서 모든 날짜 텍스트 추출 후 키워드 제거 및 공백 제거
        date_elements = driver.find_elements(By.CSS_SELECTOR, '.sc-1nljja5-0.fyJijg')

        # 키워드를 반복적으로 제거하는 코드
        date_text = ''.join([
            text.replace(" ", "") for text in [
                reduce(lambda t, keyword: t.replace(keyword, ""), remove_keywords, element.text) 
                for element in date_elements
            ]
        ])

        # 날짜 데이터를 이어 붙여 파일 이름 생성
        filename = get_unique_filename(f"{date_text}_1P", 'xlsx')

        # 엑셀 파일로 저장
        df.to_excel(filename, index=False)
        print(f"데이터가 '{filename}' 파일로 저장되었습니다.")

    except Exception as e:
        print(f"데이터 추출 중 문제가 발생했습니다. 에러: {e}")
        driver.quit()

def main():
    """메인 작업 흐름"""
    try:
        open_login_page()
        click_login_button()
        enter_username('jhkorea111')
        wait_for_user_login()
        wait_for_ad_management_menu()
        wait_for_second_button_click()
        select_value_20()
        click_sorting_icon_and_wait_for_loading()
        extract_data_to_excel()

    finally:
        driver.quit()

# 메인 함수 실행
if __name__ == "__main__":
    main()