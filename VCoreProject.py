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

        # 7. 블로그보기 링크 추출
        rows = driver.find_elements(By.CSS_SELECTOR, "table.table tbody tr")
        blog_links = []
        for row in rows:
            link_element = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) a")
            blog_link = link_element.get_attribute("href")
            blog_links.append(blog_link)

        # 8. 추출한 블로그 링크를 로그에 표시
        log_output = "\n".join(blog_links)
        log_text.delete(1.0, tk.END)
        log_text.insert(tk.END, log_output)

        # 9. 블로그 진단 페이지로 이동
        driver.get("https://lablog.co.kr/dashboard")
        print("블로그 진단 페이지로 이동 중...(확인 전)")

        # 10. 현재 URL 확인
        try:
            WebDriverWait(driver, 10).until(EC.url_to_be("https://lablog.co.kr/dashboard"))
            current_url = driver.current_url
            
            # URL이 대시보드로 변경되었는지 확인
            if "dashboard" in current_url:
                print("자동 로그인 상태입니다.")
            else:
                print("로그인 상태가 아닙니다. 로그인 절차를 진행합니다.")
                # 11. 구글 로그인 버튼 클릭 (네 번째 버튼)
                driver.get("https://lablog.co.kr/")  # 로그인 페이지로 돌아가기
                
                # 페이지가 로드될 때까지 대기
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'MuiButton-root')]")))
                try:
                    google_login_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "(//button[contains(@class, 'MuiButton-root')])[4]"))
                    )
                    google_login_button.click()
                    print("구글 로그인 버튼 클릭 완료.")
                except Exception as e:
                    print(f"구글 로그인 버튼 클릭 실패: {e}")

                # 12. 새로운 창으로 전환
                WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                original_window = driver.current_window_handle
                
                for handle in driver.window_handles:
                    if handle != original_window:
                        driver.switch_to.window(handle)

                # 13. 구글 로그인 페이지에서 ID 입력
                time.sleep(2)
                email_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "identifier"))
                )
                email_input.send_keys('marketing11111111111@gmail.com')  # 이메일 입력
                email_input.send_keys(Keys.RETURN)

                # 14. 비밀번호 입력을 위한 대기 (사용자가 직접 입력)
                print("비밀번호를 입력하세요. 브라우저는 닫히지 않으니 입력 후 진행하세요.")

                # 15. 로그인 후 블로그 진단 페이지로 이동
                WebDriverWait(driver, 10).until(EC.url_to_be("https://lablog.co.kr/dashboard"))

        except Exception as e:
            print(f"로그인 상태 확인 중 오류 발생: {e}")

        # 16. 블로그 진단 버튼 클릭
        blog_diagnose_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/blog/blogDiagnose')]"))
        )
        driver.execute_script("arguments[0].click();", blog_diagnose_button)  # JavaScript로 클릭
        
        # 블로그 진단 페이지로 이동 후 대기
        print("블로그 진단 페이지로 이동 중...")
        time.sleep(5)  # 5초 대기

        # 17. 블로그 링크 입력 및 검색
        for index, blog_link in enumerate(blog_links):
            try:
                if blog_link.startswith("https://"):  # https로 시작하는지 확인
                    # 입력 필드에 URL 입력
                    input_field = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
                    )
                    
                    # 기존 값을 완전히 지우기
                    driver.execute_script("arguments[0].value = '';", input_field)  # JavaScript로 값 제거

                    # URL 입력
                    input_field.send_keys(blog_link)  # 블로그 링크 입력
                    print(f"'{blog_link}' 입력 중...")

                    # 검색 버튼 클릭
                    search_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'MuiButton-containedPrimary')]"))
                    )
                    search_button.click()

                    # 페이지 이동 대기
                    print(f"'{blog_link}' 검색 중...")
                    WebDriverWait(driver, 10).until(EC.url_changes(driver.current_url))  # URL 변경 대기

                    # 페이지 이동 후 3초 대기
                    print("페이지 이동 완료. 3초 대기 중...")
                    time.sleep(3)

            except Exception as e:
                print(f"오류 발생: {e} - 인덱스 {index}의 블로그 링크: {blog_link}")

        print("모든 블로그 링크에 대해 검색을 완료했습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")

    finally:
        driver.quit()

def on_submit():
    url = url_entry.get()
    # 크롤링 작업을 별도의 스레드에서 실행
    threading.Thread(target=crawl_and_login, args=(url,)).start()

# GUI 설정
root = tk.Tk()
root.title("크롤링 프로그램")
root.geometry("500x400")
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
log_label = tk.Label(root, text="추출한 블로그 링크:", font=label_font, bg="#f0f0f0", fg="#333")
log_label.pack(pady=5)

log_text = scrolledtext.ScrolledText(root, width=58, height=10, font=label_font)
log_text.pack(pady=5)

# GUI 실행
root.mainloop()
