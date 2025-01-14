import requests
from bs4 import BeautifulSoup

# URL 설정
login_page_url = "https://accounts.interpark.com/authorize/ticket-pc"
login_url = "https://accounts.interpark.com/api/authentication/login"
goods_info_url = "https://ticket.interpark.com/Ticket/Goods/GoodsInfo.asp?GoodsCode=25000340"
seat_selection_url = "https://poticket.interpark.com/Book/BookMain.asp?GoodsCode=25000340"

# 세션 생성
session = requests.Session()

# 1. CSRF 토큰 추출 (헤더 추가)
login_page_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Referer": "https://ticket.interpark.com",
}

response = session.get(login_page_url, headers=login_page_headers)
soup = BeautifulSoup(response.text, 'html.parser')

# CSRF 토큰 추출
csrf_meta = soup.find("meta", {"name": "x-csrf-token"})
csrf_token = csrf_meta["content"] if csrf_meta else None

if not csrf_token:
    print("CSRF 토큰을 찾을 수 없습니다.")
    print(f"응답 내용: {response.text}")  # 디버깅용
    exit()

print(f"추출된 CSRF 토큰: {csrf_token}")

# 2. 로그인 요청
login_data = {
    "username": "qjawns2589",
    "password": "irelia13!",
    "client_id": "ticket-pc",
    "remember_me": False,
    "pc_id": "173640355794659723"
}

login_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Content-Type": "application/json",
    "Origin": "https://accounts.interpark.com",
    "Referer": login_page_url,
    "x-csrf-token": csrf_token
}

login_response = session.post(login_url, json=login_data, headers=login_headers)

# 로그인 성공 여부 확인
if login_response.status_code == 200 and "id_token" in login_response.text:
    print("로그인 성공")
    login_cookies = login_response.cookies.get_dict()
    print(f"로그인 쿠키: {login_cookies}")
    # 세션에 쿠키 설정
    for key, value in login_cookies.items():
        session.cookies.set(key, value)
else:
    print("로그인 실패")
    print(f"응답 내용: {login_response.text}")
    exit()

# 3. 공연 정보 페이지 접근
goods_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Referer": login_page_url
}
goods_response = session.get(goods_info_url, headers=goods_headers)
if goods_response.status_code == 200:
    print("공연 정보 페이지 접근 성공")
else:
    print("공연 정보 페이지 접근 실패")
    exit()

# 4. 좌석 선택 페이지 접근
seat_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Referer": goods_info_url
}
seat_response = session.get(seat_selection_url, headers=seat_headers)

# 좌석 선택 페이지 확인
if "좌석" in seat_response.text or "seat" in seat_response.text:
    print("좌석 선택 페이지 접근 성공")
    soup = BeautifulSoup(seat_response.text, 'html.parser')
    print(soup.prettify())
else:
    print("좌석 선택 페이지 접근 실패")
    print(f"응답 내용: {seat_response.text}")
