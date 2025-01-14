import time
from tkinter import Tk, Label, Button, Checkbutton, BooleanVar
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import schedule
import threading

# 사용자 데이터 디렉토리 경로 설정 (기존 로그인 세션 유지)
USER_DATA_DIR = r"C:\Interpark"

# ChromeDriver 경로 설정
driver_path = r"C:\Program Files\SeleniumBasic\chromedriver.exe"
service = Service(driver_path)

# Selenium WebDriver 설정
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=" + USER_DATA_DIR)  # 사용자 데이터 디렉토리 설정
options.add_argument("--profile-directory=Profile 1")  # 새 프로파일 생성
options.add_argument("disable-blink-features=AutomationControlled")  # 봇 탐지 방지

# Referer 헤더 설정 함수
def set_referer_header(driver):
    driver.execute_cdp_cmd(
        "Network.setExtraHTTPHeaders",
        {"headers": {"Referer": "https://ticket.interpark.com/"}}
    )

# 로그인 함수
def perform_login(driver):
    try:
        print("로그인 페이지 감지됨, 로그인 시도 중...")
        id_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "username"))
        )
        id_input.send_keys("qjawns2589")

        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "password"))
        )
        password_input.send_keys("irelia13!")

        driver.find_element(By.CSS_SELECTOR, "button.button_btnStyle__SEYzh").click()
        print("로그인 완료")
    except Exception as e:
        print(f"로그인 중 오류 발생: {e}")

# 예매 스크립트
def ticket_booking():
    driver = webdriver.Chrome(service=service, options=options)
    try:
        booking_url = "https://tickets.interpark.com/goods/25000084#"
        driver.get(booking_url)
        print("예매 페이지 접근 완료")

        # 로그인 페이지로 리다이렉트되었는지 확인
        if "accounts.interpark.com" in driver.current_url:
            perform_login(driver)
            driver.get(booking_url)
            print("예매 페이지 재접근 완료")

        # 팝업 닫기 시도
        try:
            popup_close = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "popup-prdGuide"))
            )
            driver.execute_script("arguments[0].style.display='none';", popup_close)
            print("팝업 닫기 완료")
        except Exception as e:
            print("팝업이 감지되지 않았거나 닫을 수 없습니다. 무시하고 진행합니다.")

        # 예매하기 버튼 클릭 대기 및 클릭
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.sideBtn.is-primary"))
        ).click()
        print("예매하기 버튼 클릭 완료")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        driver.quit()

# UI 생성
root = Tk()
root.title("Interpark 티켓팅")

Label(root, text="스케줄링 여부").pack(pady=10)
schedule_enabled = BooleanVar(value=False)
Checkbutton(root, text="스케줄링 사용", variable=schedule_enabled).pack(pady=5)

def start_schedule():
    if schedule_enabled.get():
        print("스케줄링 활성화: 13:59:58에 실행 예정")
        schedule.every().day.at("13:59:58").do(ticket_booking)

        def run_schedule():
            while True:
                schedule.run_pending()
                time.sleep(1)

        threading.Thread(target=run_schedule, daemon=True).start()
    else:
        print("스케줄링 비활성화, 바로 실행")
        ticket_booking()

Button(root, text="시작", command=start_schedule).pack(pady=20)
root.mainloop()
