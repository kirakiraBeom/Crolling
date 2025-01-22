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

def check_seat_availability(driver):
    try:
        # 모든 iframe 출력
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"총 {len(iframes)}개의 iframe 발견.")

        # 올바른 iframe으로 전환
        for index, iframe in enumerate(iframes):
            driver.switch_to.default_content()  # 기본 컨텍스트로 돌아가기
            driver.switch_to.frame(iframe)
            print(f"{index + 1}번째 iframe으로 전환 시도...")

            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "map[name='Map']"))
                )
                print("map 요소 발견!")
                break  # map 요소를 찾으면 해당 iframe 유지
            except:
                print("map 요소가 해당 iframe에 없습니다.")
                continue
        else:
            print("어느 iframe에서도 map 요소를 찾을 수 없습니다.")
            return

        # map 요소가 로드될 때까지 대기
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "map[name='Map']"))
        )
        print("map 요소 로드 완료.")

        # 좌석 탐색 반복
        while True:
            print("좌석 탐색 시작...")
            area_elements = driver.find_elements(By.CSS_SELECTOR, "map[name='Map'] area")
            print(f"총 {len(area_elements)}개의 좌석 영역 탐색 중...")

            for area in area_elements:
                href = area.get_attribute("href")
                title = area.get_attribute("title") or "Unknown"

                if "GetBlockSeatList" in href:
                    print(f"{title} 영역 확인 중...")

                    # 좌석 영역 클릭
                    try:
                        driver.execute_script("arguments[0].click();", area)
                        print(f"{title} 영역 클릭 완료")
                        time.sleep(2)  # 좌석 로드 대기
                    except Exception as click_error:
                        print(f"좌석 영역 클릭 중 오류 발생: {click_error}")
                        continue

                    # 활성화된 좌석 탐색
                    try:
                        seat_elements = driver.find_elements(By.CSS_SELECTOR, "span.SeatN")
                        if seat_elements:
                            for seat in seat_elements:
                                seat_title = seat.get_attribute("title")
                                print(f"활성화된 좌석 발견: {seat_title}")

                                # 예매 가능한 좌석 클릭
                                try:
                                    seat.click()
                                    print(f"좌석 선택 완료: {seat_title}")
                                    return
                                except Exception as click_error:
                                    print(f"좌석 클릭 중 오류 발생: {click_error}")
                        else:
                            print(f"{title} 영역: 활성화된 좌석 없음")
                    except Exception as seat_error:
                        print(f"좌석 탐색 중 오류 발생: {seat_error}")

                    # 이전 화면으로 복귀
                    try:
                        driver.back()
                        time.sleep(1)
                    except Exception as back_error:
                        print(f"이전 화면 복귀 중 오류 발생: {back_error}")

            print("빈 좌석이 없습니다. 5초 후 다시 탐색합니다...")
            time.sleep(5)  # 일정 시간 대기 후 다시 탐색

    except Exception as e:
        print(f"좌석 탐색 중 오류 발생: {e}")
        traceback.print_exc()

# 캡챠 입력 대기 함수
def wait_for_captcha_completion(driver):
    try:
        print("캡챠 입력 대기 중...")
        WebDriverWait(driver, 600).until(
            lambda d: "none" in d.find_element(By.ID, "divRecaptcha").get_attribute("style").lower()
        )
        print("캡챠 입력 완료. 좌석 탐색 시작!")
        check_seat_availability(driver)
    except Exception as e:
        print("캡챠 입력 확인 중 오류 발생:")
        traceback.print_exc()

# 새 창 대기 및 URL 확인 후 전환 함수
def wait_for_specific_url_and_switch(driver):
    try:
        original_window = driver.current_window_handle
        print(f"현재 창 핸들: {original_window}")

        # 새 창이 열릴 때까지 대기
        WebDriverWait(driver, 30).until(lambda d: len(d.window_handles) > 1)
        print("새 창이 감지됨. URL 확인 대기 중...")

        while True:
            for window in driver.window_handles:
                driver.switch_to.window(window)
                current_url = driver.current_url
                print(f"현재 새 창 URL: {current_url}")

                # 원하는 URL이 감지되면 전환 완료
                if "poticket.interpark.com/Book/BookMain.asp" in current_url:
                    print(f"원하는 URL로 전환 완료: {current_url}")
                    return

            # 원하는 URL이 없으면 1초 대기 후 다시 확인
            time.sleep(1)

    except Exception as e:
        print(f"새 창 전환 중 오류 발생: {e}")
        traceback.print_exc()

# 좌석 선택 페이지 대기
def wait_for_seat_selection_page(driver):
    try:
        print("좌석 선택 페이지로 전환 대기 중...")
        WebDriverWait(driver, 600).until(
            EC.url_contains("poticket.interpark.com/Book/BookMain.asp")
        )
        print("좌석 선택 페이지로 전환 완료!")
    except Exception as e:
        print("좌석 선택 페이지로 전환되지 않음. 오류 발생:")
        traceback.print_exc()

# 예매 스크립트
def ticket_booking(goods_code):
    driver = webdriver.Chrome(service=service, options=options)
    try:
        booking_url = f"https://tickets.interpark.com/goods/{goods_code}#"
        driver.get(booking_url)
        print(f"예매 페이지 접근 완료 (GoodsCode: {goods_code})")

        # 팝업 닫기 시도
        close_popup(driver)

        # 예매하기 버튼 클릭
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a.sideBtn.is-primary"))
            ).click()
            print("예매하기 버튼 클릭 완료")
        except Exception as e:
            print("예매하기 버튼 클릭 실패")
            traceback.print_exc()

        # 새 창 전환 전에 10초 대기
        print("새 창 전환 전에 10초 대기 중...")
        time.sleep(10)

        # 새 창 대기 및 URL 확인 후 전환
        print("새 창 대기 및 URL 확인 중...")
        wait_for_specific_url_and_switch(driver)

        # 좌석 탐색 시작
        check_seat_availability(driver)

    except Exception as e:
        print(f"오류 발생: {e}")
        traceback.print_exc()
    # finally:
    #     driver.quit()

# 스케줄링 시작
def start_schedule(goods_code):
    schedule_time = schedule_time_var.get()
    if not schedule_time:
        print("스케줄링 시간이 입력되지 않았습니다. 기본값 13:59:58로 설정합니다.")
        schedule_time = "13:59:58"

    if schedule_enabled.get():
        print(f"스케줄링 활성화: {schedule_time}에 실행 예정 (GoodsCode: {goods_code})")
        schedule.every().day.at(schedule_time).do(ticket_booking, goods_code=goods_code)

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

Label(root, text="스케줄링 시간 입력 (예: 13:59:58)").pack(pady=10)
schedule_time_var = StringVar()
Entry(root, textvariable=schedule_time_var).pack(pady=5)

Label(root, text="스케줄링 여부").pack(pady=10)
schedule_enabled = BooleanVar(value=False)
Checkbutton(root, text="스케줄링 사용", variable=schedule_enabled).pack(pady=5)

Button(root, text="시작", command=lambda: start_schedule(goods_code_var.get())).pack(pady=20)

root.mainloop()
