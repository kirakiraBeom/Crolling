import time
from tkinter import Tk, Label, Button, Checkbutton, BooleanVar, Entry, StringVar
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import schedule
import threading
import traceback  # 스택 트레이스 출력용

# 사용자 데이터 디렉토리 경로 설정
USER_DATA_DIR = r"C:\Interpark"
login_page_url = "https://accounts.interpark.com/authorize/ticket-pc?postProc=FULLSCREEN&origin=https%3A%2F%2Fticket.interpark.com%2FGate%2FTPLoginConfirmGate.asp%3FGroupCode%3D%26Tiki%3D%26Point%3D%26PlayDate%3D%26PlaySeq%3D%26HeartYN%3D%26TikiAutoPop%3D%26BookingBizCode%3D%26MemBizCD%3DWEBBR%26CPage%3D%26GPage%3Dhttps%253A%252F%252Ftickets.interpark.com%252Fgoods%252F25000084%2523&version=v2"

# ChromeDriver 경로 설정
driver_path = r"C:\Program Files\SeleniumBasic\chromedriver.exe"
service = Service(driver_path)

# Selenium WebDriver 설정
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=" + USER_DATA_DIR)
options.add_argument("--profile-directory=Profile 1")
options.add_argument("disable-blink-features=AutomationControlled")

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
        traceback.print_exc()  # 스택 트레이스 출력

# 팝업 닫기 함수
def close_popup(driver):
    try:
        popup_close = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "popup-prdGuide"))
        )
        driver.execute_script("arguments[0].style.display='none';", popup_close)
        print("팝업 닫기 완료")
    except Exception as e:
        print("팝업이 감지되지 않았거나 닫을 수 없습니다. 무시하고 진행합니다.")
        traceback.print_exc()  # 스택 트레이스 출력

# 예매 스크립트
def ticket_booking(goods_code):
    driver = webdriver.Chrome(service=service, options=options)
    try:
        booking_url = f"https://tickets.interpark.com/goods/{goods_code}#"
        driver.get(booking_url)
        print(f"예매 페이지 접근 완료 (GoodsCode: {goods_code})")

        # 팝업 닫기 시도
        close_popup(driver)

        # 예매하기 버튼 클릭 대기
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a.sideBtn.is-primary"))
            ).click()
            print("예매하기 버튼 클릭 완료")
        except Exception:
            print("예매하기 버튼 클릭 실패. 로그인 여부 확인 중...")

        # 현재 URL 확인 및 로그인 처리
        WebDriverWait(driver, 10).until(EC.url_contains("accounts.interpark.com"))
        current_url = driver.current_url
        print(f"현재 URL: {current_url}")

        if "accounts.interpark.com/authorize/ticket-pc" in current_url:
            perform_login(driver)
            # 팝업 닫기 시도
            close_popup(driver)

            # 로그인 후 예매하기 버튼 재클릭
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a.sideBtn.is-primary"))
            ).click()
            print("로그인 후 예매하기 버튼 클릭 완료")

        # 새 창으로 전환
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])
        print("새 창으로 전환 완료")

    except Exception as e:
        print(f"오류 발생: {e}")
    # finally:
    #     driver.quit()

# 스케줄링 시작
def start_schedule(goods_code):
    if schedule_enabled.get():
        print(f"스케줄링 활성화: 13:59:58에 실행 예정 (GoodsCode: {goods_code})")
        schedule.every().day.at("13:59:58").do(ticket_booking, goods_code=goods_code)

        def run_schedule():
            while True:
                schedule.run_pending()
                time.sleep(1)

        threading.Thread(target=run_schedule, daemon=True).start()
    else:
        print("스케줄링 비활성화, 바로 실행")
        ticket_booking(goods_code)

# UI 생성
root = Tk()
root.title("Interpark 티켓팅")

Label(root, text="GoodsCode 입력").pack(pady=10)
goods_code_var = StringVar()
Entry(root, textvariable=goods_code_var).pack(pady=5)

Label(root, text="스케줄링 여부").pack(pady=10)
schedule_enabled = BooleanVar(value=False)
Checkbutton(root, text="스케줄링 사용", variable=schedule_enabled).pack(pady=5)

Button(root, text="시작", command=lambda: start_schedule(goods_code_var.get())).pack(pady=20)

root.mainloop()
