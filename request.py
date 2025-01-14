import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time

# ChromeDriver 경로 설정
driver_path = r"C:\Program Files\SeleniumBasic\chromedriver.exe"
service = Service(driver_path)

# Selenium WebDriver 실행
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# Interpark URL 설정
login_page_url = "https://accounts.interpark.com/authorize/ticket-pc?postProc=FULLSCREEN&origin=https%3A%2F%2Fticket.interpark.com%2FGate%2FTPLoginConfirmGate.asp%3FGroupCode%3D%26Tiki%3D%26Point%3D%26PlayDate%3D%26PlaySeq%3D%26HeartYN%3D%26TikiAutoPop%3D%26BookingBizCode%3D%26MemBizCD%3DWEBBR%26CPage%3D%26GPage%3Dhttps%253A%252F%252Ftickets.interpark.com%252Fgoods%252F25000340%253F%2523&version=v2"
login_url = "https://accounts.interpark.com/api/authentication/login"
seat_selection_url = "https://poticket.interpark.com/Book/BookMain.asp?GoodsCode=25000340"

# 1. Selenium으로 로그인 페이지 로드
driver.get(login_page_url)
time.sleep(2)  # 페이지 로딩 대기

# CSRF 토큰 추출
soup = BeautifulSoup(driver.page_source, 'html.parser')
csrf_meta = soup.find("meta", {"name": "x-csrf-token"})
csrf_token = csrf_meta["content"] if csrf_meta else None

if not csrf_token:
    print("CSRF 토큰을 찾을 수 없습니다.")
    driver.quit()
    exit()

print(f"추출된 CSRF 토큰: {csrf_token}")
 
# PCID 추출
cookies = driver.get_cookies()
pc_id = None
for cookie in cookies:
    if cookie['name'] == 'pcid':
        pc_id = cookie['value']
        break

if not pc_id:
    print("PCID를 찾을 수 없습니다.")
    driver.quit()
    exit()

print(f"추출된 PCID: {pc_id}")

# 2. 로그인 요청
session = requests.Session()
login_data = {
    "username": "qjawns2589",
    "password": "irelia13!",
    "client_id": "ticket-pc",
    "remember_me": False,
    "pc_id": pc_id
}

login_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Content-Type": "application/json",
    "Origin": "https://accounts.interpark.com",
    "Referer": login_page_url,
    "x-csrf-token": csrf_token,
    "sec-fetch-site": "same-origin",
    "sec-fetch-mode": "cors",
    "sec-fetch-dest": "empty",
    "sec-ch-ua": '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
}

login_response = session.post(login_url, json=login_data, headers=login_headers)

if login_response.status_code == 200 and "id_token" in login_response.text:
    print("로그인 성공")
    login_cookies = session.cookies.get_dict()
    print(f"로그인 쿠키: {login_cookies}")
else:
    print("로그인 실패")
    print(f"응답 내용: {login_response.text}")
    driver.quit()
    exit()

# 3. Selenium에 로그인 세션 쿠키 설정
driver.delete_all_cookies()
for key, value in login_cookies.items():
    driver.add_cookie({"name": key, "value": value, "domain": ".interpark.com"})

# 4. 좌석 선택 페이지로 이동
driver.get(seat_selection_url)
print("좌석 선택 페이지 접근 시도 중...")

time.sleep(5)
if "좌석" in driver.page_source:
    print("좌석 선택 페이지 접근 성공")
else:
    print("좌석 선택 페이지 접근 실패")

driver.quit()
