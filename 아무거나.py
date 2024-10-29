import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import pandas as pd  # 데이터프레임 처리
from openpyxl import Workbook  # 엑셀 저장


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
    while True:
        try:
            login_button = WebDriverWait(driver, 180).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@type="submit"]'))
            )
            login_button.click()
            break  # 성공적으로 클릭하면 루프 종료
        except TimeoutException:
            print("로그인 버튼을 찾지 못했습니다. 다시 시도합니다...")
            time.sleep(5)  # 5초 대기 후 재시도

def wait_for_ad_management_menu():
    """광고 관리 메뉴가 나타날 때까지 대기 후 클릭"""
    ad_management_menu = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//a[text()="광고관리"]'))
    )
    driver.execute_script("arguments[0].click();", ad_management_menu)
    time.sleep(3)  # 페이지 로드 대기
    
def wait_for_second_button_click():
    """사용자가 두 번째 버튼을 클릭할 때까지 대기"""
    while True:
        try:
            second_button = WebDriverWait(driver, 1800).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.ant-btn-primary.DateRangePickerPanel__DateRangePickerPanelButton-sc-1lnvml5-3.ekYdWg'))
            )
            print("두 번째 버튼이 나타났습니다. 클릭을 기다리는 중...")

            WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ant-btn-primary.DateRangePickerPanel__DateRangePickerPanelButton-sc-1lnvml5-3.ekYdWg')))
            WebDriverWait(driver, 180).until_not(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '.ant-btn-primary.DateRangePickerPanel__DateRangePickerPanelButton-sc-1lnvml5-3.ekYdWg'))
            )
            print("사용자가 두 번째 버튼을 클릭했습니다.")
            time.sleep(2)
            break  # 클릭 후 루프 종료

        except Exception as e:
            print(f"버튼을 감지할 수 없습니다. 다시 시도합니다... 에러: {e}")
            time.sleep(5)  # 5초 대기 후 재시도

def extract_date_range():
    """DateRangePickerText__DateRangePickerTextWrapper-sc-12l6wzj-0 클래스 안에 있는 모든 span 데이터를 추출"""
    try:
        # 제거할 문자열 목록 (필요에 따라 여기에 추가)
        remove_strings = ["지난달:", "지난주:", "최근30일:"]

        # 해당 클래스 내부의 모든 span 태그를 찾음
        date_elements = driver.find_elements(By.CSS_SELECTOR, '.DateRangePickerText__DateRangePickerTextWrapper-sc-12l6wzj-0.fcdYIv span')
        
        # 각 span 태그의 텍스트를 가져와서 리스트로 저장하고 제거할 문자열들을 제거
        date_texts = []
        for element in date_elements:
            text = element.text.replace(" ", "")  # 공백 제거
            for remove_string in remove_strings:
                text = text.replace(remove_string, "")  # 제거할 문자열 반복적으로 제거
            date_texts.append(text)

        # 추출된 모든 날짜 텍스트를 ''로 연결하여 하나의 문자열로 만듦
        date_range_text = "".join(date_texts)

        # 날짜 범위 출력 확인
        print(f"추출된 날짜 범위: {date_range_text}")
        
        # 파일 이름으로 사용될 날짜 범위 반환
        return date_range_text

    except Exception as e:
        print(f"날짜 범위 추출 중 문제가 발생했습니다. 에러: {e}")
        driver.quit()
        return None

def select_value_20():
    """페이지 당 20개 옵션 선택 후 3초 대기"""
    try:
        select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//select[@aria-label="rows per page"]'))
        )
        select = Select(select_element)
        select.select_by_value('20')  # value가 20인 옵션 선택
        print("옵션 20개 선택 완료.")
        time.sleep(3)

    except TimeoutException:
        print("옵션 선택 실패.")

def click_campaign_with_keyword(keyword):
    """캠페인 이름에 특정 키워드가 포함된 캠페인을 클릭하고 스크롤하여 해당 요소 표시"""
    try:
        # content-main-area 클래스 내부에서 rt-tbody 클래스를 포함한 요소 찾기
        content_main_area = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'content-main-area'))
        )

        # JavaScript를 사용하여 rt-tbody 요소 찾기
        tbody = driver.execute_script("return document.querySelector('.content-main-area .rt-tbody');")
        
        if tbody is None:
            print("rt-tbody 요소를 찾을 수 없습니다.")
            return False

        last_height = driver.execute_script("return arguments[0].scrollHeight;", tbody)
        
        while True:
            rows = tbody.find_elements(By.CLASS_NAME, 'rt-tr-group')
            for row in rows:
                # rt-tr-group 클래스 안에서 a 태그를 찾고 텍스트가 "추출용"인지 확인
                link_element = row.find_element(By.TAG_NAME, 'a')
                if keyword in link_element.text:
                    print(f"'{link_element.text}' 캠페인 발견. 클릭 중...")
                    
                    # 자바스크립트를 사용하여 안전하게 클릭
                    driver.execute_script("arguments[0].click();", link_element)
                    time.sleep(5)  # 페이지 로딩 시간 충분히 대기
                    return True

            # 스크롤을 내려 더 많은 캠페인 요소를 로드하도록 함
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", tbody)
            time.sleep(2)  # 로딩 시간 대기

            # 더 이상 스크롤할 수 없으면 반복문 탈출
            new_height = driver.execute_script("return arguments[0].scrollHeight;", tbody)
            if new_height == last_height:
                print(f"캠페인 이름에 '{keyword}' 포함된 캠페인을 찾을 수 없습니다.")
                break
            last_height = new_height

        return False

    except Exception as e:
        print(f"캠페인 클릭 중 문제가 발생했습니다. 에러: {e}")
        driver.quit()

def click_dashboard_title():
    """rt-tbody 클래스 내부의 dashboard-title-wrapper 클래스의 a 태그 클릭"""
    try:
        # Table__StyledTable 클래스 내부의 rt-tbody 클래스를 포함한 요소 찾기
        table_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'Table__StyledTable-sc-nenwbp-0'))
        )

        # JavaScript를 사용하여 rt-tbody 요소 찾기
        tbody = table_element.find_element(By.CLASS_NAME, 'rt-tbody')
        
        if tbody is None:
            print("rt-tbody 요소를 찾을 수 없습니다.")
            return False

        # rt-tbody 내부의 dashboard-title-wrapper 클래스 찾기
        dashboard_title_wrapper = tbody.find_element(By.CLASS_NAME, 'dashboard-title-wrapper')
        
        # dashboard-title-wrapper 내부의 a 태그 클릭
        dashboard_title_link = dashboard_title_wrapper.find_element(By.TAG_NAME, 'a')
        
        # 자바스크립트를 사용하여 안전하게 클릭
        driver.execute_script("arguments[0].click();", dashboard_title_link)
        time.sleep(5)  # 페이지 로딩 시간 충분히 대기
        print("대시보드 제목 클릭 완료.")
        return True

    except Exception as e:
        print(f"대시보드 제목 클릭 중 문제가 발생했습니다. 에러: {e}")
        driver.quit()
        return False
    
def select_value_50():
    """페이지 당 50개 옵션 선택 후 3초 대기"""
    try:
        select_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//select[@aria-label="rows per page"]'))
        )
        select = Select(select_element)
        select.select_by_value('50')  # value가 50인 옵션 선택
        print("옵션 50개 선택 완료.")
        time.sleep(3)

    except TimeoutException:
        print("옵션 선택 실패.")

def get_unique_filename(filename):
    """파일 이름이 중복될 경우 _1, _2, ... 형식으로 새로운 파일 이름 생성"""
    base, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename

    while os.path.exists(new_filename):
        new_filename = f"{base}_{counter}{extension}"
        counter += 1

    return new_filename

def extract_table_data_and_save_to_excel(filename):
    """테이블 데이터 추출 및 엑셀 파일로 저장"""
    while True:
        try:
            # Table__StyledTable 클래스 내부의 rt-tbody 클래스를 포함한 요소 찾기
            table_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'Table__StyledTable-sc-nenwbp-0'))
            )

            # JavaScript를 사용하여 rt-tbody 요소 찾기
            tbody = table_element.find_element(By.CLASS_NAME, 'rt-tbody')
            
            if tbody is None:
                print("rt-tbody 요소를 찾을 수 없습니다.")
                return False

            # 모든 rt-tr-group 클래스 내부 요소 탐색
            rows = tbody.find_elements(By.CLASS_NAME, 'rt-tr-group')
            data = []

            # 각 행에서 데이터 추출
            for row in rows:
                product_name = row.find_element(By.CLASS_NAME, 'useProductsTableColumns__ProductNameWrapper-sc-139x206-0').text.replace("지난달:", "")
                ad_spend = row.find_elements(By.CLASS_NAME, 'rt-td')[4].text
                conversion_sales = row.find_elements(By.CLASS_NAME, 'rt-td')[5].text
                impressions = row.find_elements(By.CLASS_NAME, 'rt-td')[6].text
                clicks = row.find_elements(By.CLASS_NAME, 'rt-td')[7].text
                click_through_rate = row.find_elements(By.CLASS_NAME, 'rt-td')[8].text
                conversion_sales_count = row.find_elements(By.CLASS_NAME, 'rt-td')[9].text
                conversion_rate = row.find_elements(By.CLASS_NAME, 'rt-td')[10].text
                roas = row.find_elements(By.CLASS_NAME, 'rt-td')[11].text
                conversion_orders = row.find_elements(By.CLASS_NAME, 'rt-td')[12].text

                data.append([product_name, ad_spend, conversion_sales, impressions, clicks, click_through_rate, conversion_sales_count, conversion_rate, roas, conversion_orders])

            # 데이터프레임 생성
            columns = [
                "상품명", "집행광고비", "광고전환매출", "노출수", "클릭수", 
                "클릭률", "광고 전환 판매수", "전환율", "광고수익률", "광고 전환 주문수"
            ]
            df = pd.DataFrame(data, columns=columns)
            
            # 파일 이름 중복 확인 및 저장
            unique_filename = get_unique_filename(filename)
            df.to_excel(unique_filename, index=False)
            print(f"데이터가 '{unique_filename}' 파일에 저장되었습니다.")
            break  # 성공적으로 저장하면 루프 종료
            
        except Exception as e:
            print(f"데이터 추출 및 저장 중 문제가 발생했습니다. 에러: {e}")
            time.sleep(5)  # 5초 대기 후 재시도

def main():
    """메인 작업 흐름"""
    try:
        open_login_page()
        click_login_button()
        enter_username('jhkorea111')
        wait_for_user_login()
        wait_for_ad_management_menu()
        wait_for_second_button_click()

        # 날짜 범위를 추출하여 파일 이름에 반영
        date_range = extract_date_range()
        if not date_range:
            print("날짜 범위 추출에 실패하여 작업을 중단합니다.")
            return

        select_value_20()
        click_campaign_with_keyword("추출용")
        click_dashboard_title()
        select_value_50()

        # 테이블 데이터 추출 및 엑셀 파일로 저장 (파일 이름에 날짜 범위 포함)
        extract_table_data_and_save_to_excel(f"{date_range}_추출용.xlsx")

    except Exception as e:
        print(f"메인 작업 중 오류 발생: {e}")

    finally:
        driver.quit()

# 메인 함수 실행
if __name__ == "__main__":
    main()