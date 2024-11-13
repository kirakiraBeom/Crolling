# https://xn--939au0g4vj8sq.net/rq/?k=MTU1NDYwNg==
# https://xn--939au0g4vj8sq.net/rq/?k=MTU2NzYwMA==

import tkinter as tk
from tkinter import font, scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import threading

# 크롬 드라이버 설정
driver_path = "C:\\Program Files\\SeleniumBasic\\chromedriver.exe"
service = Service(driver_path)
chrome_options = Options()

# 데이터 기억 옵션 추가
chrome_options.add_argument("user-data-dir=C:\\VCP")  # 사용자 데이터 디렉토리 설정
chrome_options.add_argument("disable-blink-features=AutomationControlled")
# chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36")

def crawl_and_login(url):
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        # 1. 사용자가 입력한 URL로 이동
        driver.get(url)
        print("URL로 이동 중...")

        # 2. 페이지 로딩 대기
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="password"]')))
        print("비밀번호 입력 필드가 로드되었습니다.")

        # 3. 비밀번호 입력 필드 찾기
        text_input = driver.find_element(By.CSS_SELECTOR, 'input[type="password"]')

        # 4. '3605' 입력
        text_input.send_keys('3605')  # 비밀번호 입력
        print("비밀번호 입력 완료.")

        # 5. 비밀번호 제출
        text_input.send_keys(Keys.RETURN)
        print("비밀번호 제출 중...")

        # 6. 페이지 로딩 대기
        time.sleep(5)

        # 7. 블로그 링크, 주소, 연락처, 블로그 ID 추출
        rows = driver.find_elements(By.CSS_SELECTOR, "table.table tbody tr")
        blog_links = []
        blog_ids = []
        addresses = []
        contacts = []

        for row in rows:
            # 블로그 링크 추출
            link_element = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) a")
            blog_link = link_element.get_attribute("href")
            blog_links.append(blog_link)

            # 블로그 ID 추출
            blog_id = blog_link.split("/")[-1]
            blog_ids.append(blog_id)

            # # 주소와 연락처 추출
            # address = row.find_element(By.CSS_SELECTOR, "td:nth-child(5)").text
            # contact = row.find_element(By.CSS_SELECTOR, "td:nth-child(4)").text
            # addresses.append(address)
            # contacts.append(contact)

        # 8. 추출한 블로그 링크를 로그에 표시
        log_output = "\n".join(blog_ids)
        log_text.delete(1.0, tk.END)
        log_text.insert(tk.END, log_output)

        # 주소, 연락처, 블로그 ID는 기억만 해둠
        print("주소, 연락처, 블로그 ID가 기억되었습니다.")
        
    finally:
        driver.quit()

def start_research():
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    try:
        # 1. 연구 사이트로 이동
        driver.get("https://lablog.co.kr/")
        print("연구 사이트로 이동 중...")
        
        # 2. 2초 대기
        time.sleep(2)
        
        # 3. 세 번째 버튼 클릭
        elements = driver.find_elements(By.CSS_SELECTOR, ".MuiGrid-root.MuiGrid-item.MuiGrid-grid-xs-4.css-1udb513")
        if len(elements) >= 3:
            elements[2].click()  # 세 번째 요소 클릭
            print("세 번째 버튼 클릭 완료.")
        else:
            print("요소를 찾을 수 없습니다.")
            # "블로그 진단" 클릭 단계로 점프
            goto_blog_diagnosis(driver)
            return

        # 4. 새 창이 열릴 때까지 대기
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        original_window = driver.current_window_handle

        # 5. 새 창으로 전환
        for handle in driver.window_handles:
            if handle != original_window:
                driver.switch_to.window(handle)
                break
        else:
            print("새 창이 열리지 않았습니다.")
            return  # 새 창이 열리지 않으면 종료

        print("새 창으로 전환 완료.")

        # 6. 로그인 페이지에서 이메일 입력 대기 및 입력
        time.sleep(2)  # 추가 대기 (필요 시 제거)
        email_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "identifier"))
        )
        email_input.send_keys('marketing11111111111@gmail.com')  # 이메일 입력
        email_input.send_keys(Keys.RETURN)
        print("이메일 입력 및 제출 완료.")

        # 7. 새 창이 닫힐 때까지 대기
        WebDriverWait(driver, 20).until(EC.number_of_windows_to_be(1))
        driver.switch_to.window(original_window)  # 원래 창으로 전환
        print("새 창이 닫혔습니다. 원래 창으로 돌아왔습니다.")
        time.sleep(2)

        # 8. "블로그 진단" 클릭 단계로 이동
        goto_blog_diagnosis(driver)
    
    finally:
        pass

def goto_blog_diagnosis(driver):
    # "다시보지않기" 버튼이 있으면 클릭
    try:
        dont_show_again_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[text()='다시보지않기']"))
        )
        dont_show_again_button.click()
        print("다시보지않기 버튼 클릭 완료.")
    except Exception:
        print("다시보지않기 버튼이 없습니다. 진행합니다.")

    # "블로그 진단" 텍스트를 찾아 클릭
    blog_diagnosis_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[text()='블로그 진단']"))
    )
    blog_diagnosis_button.click()
    print("블로그 진단 버튼 클릭 완료.")

    # 입력 필드의 value 값 초기화 (클릭 후 백스페이스 5초)
    input_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiInputBase-input.MuiOutlinedInput-input.MuiInputBase-inputSizeSmall.css-2ozrrz"))
    )
    input_field.click()  # 입력 필드 클릭
    print("입력 필드 클릭 완료.")
    
    # 2초 동안 백스페이스 키를 누름
    start_time = time.time()
    while time.time() - start_time < 2:
        input_field.send_keys(Keys.BACKSPACE)
    print("입력 필드 초기화 완료 (백스페이스 2초).")

    # 로그에 저장된 항목을 차례대로 입력
    log_items = log_text.get("1.0", tk.END).strip().split("\n")
    for item in log_items:
        input_field.click()  # 입력 필드 클릭
        print("입력 필드 클릭 완료.")

        # 3초 동안 백스페이스 키를 누름
        start_time = time.time()
        while time.time() - start_time < 3:
            input_field.send_keys(Keys.BACKSPACE)
        print("입력 필드 초기화 완료 (백스페이스 3초).")

        # 항목 입력
        input_field.send_keys(item)
        print(f"'{item}' 입력 완료.")

        # 지정된 버튼 클릭
        action_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".MuiButtonBase-root.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeMedium.MuiButton-containedSizeMedium.MuiButton-fullWidth.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeMedium.MuiButton-containedSizeMedium.MuiButton-fullWidth.css-9nd1ol"))
        )
        action_button.click()
        print("액션 버튼 클릭 완료.")

        # 지정된 클래스가 사라질 때까지 대기
        WebDriverWait(driver, 30).until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".MuiLinearProgress-root.MuiLinearProgress-colorSecondary.MuiLinearProgress-determinate.css-1hbgb9z"))
        )
        print("지정된 클래스가 없어졌습니다. 루프를 다시 시작합니다.")

        # 5초 대기
        time.sleep(5)

def on_submit():
    url = url_entry.get()
    # 크롤링 작업을 별도의 스레드에서 실행
    threading.Thread(target=crawl_and_login, args=(url,)).start()

# GUI 설정
root = tk.Tk()
root.title("크롤링 프로그램")
root.geometry("550x550")  # 높이를 늘려 버튼 추가 공간 확보
root.configure(bg="#f0f0f0")

# 폰트 설정
title_font = font.Font(family="Helvetica", size=16, weight="bold")
label_font = font.Font(family="Helvetica", size=12)
button_font = font.Font(family="Helvetica", size=12, weight="bold")

# 제목 레이블
title_label = tk.Label(root, text="크롤링 프로그램", font=title_font, bg="#f0f0f0", fg="#333")
title_label.pack(pady=20)

# URL 입력란
url_label = tk.Label(root, text="접속할 링크를 입력하세요:", font=label_font, bg="#f0f0f0", fg="#333")
url_label.pack(pady=5)

url_entry = tk.Entry(root, width=40, font=label_font)
url_entry.pack(pady=5)

# 제출 버튼
submit_button = tk.Button(root, text="제출", command=on_submit, font=button_font, bg="#4CAF50", fg="white", padx=10, pady=5)
submit_button.pack(pady=20)

# 로그 출력 영역
log_label = tk.Label(root, text="추출한 블로그ID:", font=label_font, bg="#f0f0f0", fg="#333")
log_label.pack(pady=5)

log_text = scrolledtext.ScrolledText(root, width=58, height=10, font=label_font)
log_text.pack(pady=5)

# 연구시작 버튼 추가
research_button = tk.Button(root, text="연구시작", command=start_research, font=button_font, bg="#FF5733", fg="white", padx=10, pady=5)
research_button.pack(pady=20)

# GUI 실행
root.mainloop()
