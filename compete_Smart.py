import json
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# JSON 파일 읽기
with open('selector.json', 'r', encoding='utf-8') as file:
    selectors = json.load(file)

# 웹 드라이버 초기화
options = Options()
options.add_argument("--headless")  # 브라우저를 표시하지 않음
options.add_argument("--ignore-certificate-errors")  # SSL 인증서 오류 무시

# Chrome WebDriver 설정
driver_path = "C:\\Program Files\\SeleniumBasic\\chromedriver.exe"
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# 결과를 저장할 리스트
results = []

def crawl_homepage(homepage_name, url, a_selector, b_selector, group_type):
    try:
        driver.get(url)
        if group_type == 'A':
            buttons = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, a_selector))
            )
        else:
            buttons = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, b_selector))
            )
        return {"Homepage": homepage_name, "Button Count": len(buttons)}
    except NoSuchElementException:
        print(f"{homepage_name}에서 버튼을 찾을 수 없습니다.")
        return {"Homepage": homepage_name, "Button Count": 0}
    except Exception as e:
        print(f"{homepage_name} 크롤링 중 오류 발생: {e}")
        return {"Homepage": homepage_name, "Button Count": 0}

# A 그룹: smartstore로 시작하는 경쟁사
a_group = [
    {"name": "유투카", "url": selectors['유투카']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "세이보링", "url": selectors['세이보링']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "안녕하십니카", "url": selectors['안녕하십니카']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "킨톤", "url": selectors['킨톤']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "이지스오토랩", "url": selectors['이지스오토랩']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "기가차스토어", "url": selectors['기가차스토어']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "테스트드라이브(카마루)", "url": selectors['테스트드라이브']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "트니르", "url": selectors['트니르']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "더레이즈", "url": selectors['더레이즈']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "발상코퍼레이션", "url": selectors['발상코퍼레이션']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
    {"name": "데코M(엠노블)", "url": selectors['데코M']['url'], "a_selector": "//div[@class='ZjAm7j87Us']/a/span", "b_selector": None},
]

# B 그룹: brand로 시작하는 경쟁사
b_group = [
    {"name": "메이튼", "url": selectors['메이튼']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "벤딕트", "url": selectors['벤딕트']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "케이엠모터스", "url": selectors['케이엠모터스']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "지엠지모터스(꾸미자닷컴)", "url": selectors['지엠지모터스']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "본투로드", "url": selectors['본투로드']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "가온", "url": selectors['가온']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "아임반", "url": selectors['아임반']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "돌비웨이", "url": selectors['돌비웨이']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "카멜레온360", "url": selectors['카멜레온360']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
    {"name": "주파집", "url": selectors['주파집']['url'], "a_selector": None, "b_selector": "//div[@class='_3_Glsr7yy3']//ul/li/button/span"},
]

# A 그룹 크롤링
for competitor in a_group:
    result = crawl_homepage(competitor["name"], competitor["url"], competitor["a_selector"], competitor["b_selector"], 'A')
    results.append(result)

# B 그룹 크롤링
for competitor in b_group:
    result = crawl_homepage(competitor["name"], competitor["url"], competitor["a_selector"], competitor["b_selector"], 'B')
    results.append(result)

# 웹 드라이버 종료
driver.quit()

# CSV 파일로 저장
with open('button_counts.csv', 'w', newline='') as csvfile:
    fieldnames = ['Homepage', 'Button Count']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

    writer.writeheader()
    for result in results:
        writer.writerow(result)

print("버튼 개수가 'button_counts.csv' 파일에 저장되었습니다.")
