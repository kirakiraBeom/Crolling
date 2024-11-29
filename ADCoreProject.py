import requests
import time
import hmac
import hashlib
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import urllib.parse


# 네이버 API 설정
NAVER_CLIENT_KEY = '0100000000258e1a9e2d626b6e10c80471c41a9834842b9bf3303ce02b6f2de748ae513a1c'
NAVER_CLIENT_SECRET = 'AQAAAAAljhqeLWJrbhDIBHHEGpg0+/lKFc+RcMcyGJUgVScNeg=='
NAVER_CUSTOMER_ID = '2902933'
BASE_URL = 'https://api.searchad.naver.com'

# Google Sheets API 설정
GOOGLE_SHEET_ID = '15q1Gqr3iyQM8DM6rjmPWyYO92r-1jo5ESV4DrLltbS4'
JSON_KEY_FILE = r'C:\Crolling\json\google_sheets_key.json'

# 요청 헤더 생성 함수
def get_header(method, uri):
    timestamp = str(int(time.time() * 1000))
    message = f"{timestamp}.{method}.{uri}"
    signature = hmac.new(
        NAVER_CLIENT_SECRET.encode('utf-8'),
        message.encode('utf-8'),
        hashlib.sha256
    ).digest()
    signature = base64.b64encode(signature).decode('utf-8')

    print(f"[DEBUG] Signature Message: {message}")  # 서명 생성용 메시지 출력
    print(f"[DEBUG] X-Signature: {signature}")  # 생성된 서명 출력

    return {
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Timestamp': timestamp,
        'X-API-KEY': NAVER_CLIENT_KEY,
        'X-Customer': NAVER_CUSTOMER_ID,
        'X-Signature': signature
    }

def fetch_campaigns():
    uri = "/ncc/campaigns"
    url = f"{BASE_URL}{uri}"
    headers = get_header("GET", uri)

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    campaigns = response.json()

    # 활성화된 캠페인만 반환
    return [campaign for campaign in campaigns if campaign.get('status') == 'ELIGIBLE']


def fetch_adgroups(campaign_id):
    uri = "/ncc/adgroups"
    url = f"{BASE_URL}{uri}"
    headers = get_header("GET", uri)
    params = {"nccCampaignId": campaign_id}

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    adgroups = response.json()

    # 활성화된 광고그룹만 반환
    return [adgroup for adgroup in adgroups if adgroup.get('status') == 'ELIGIBLE']

# 광고 그룹 성과 가져오기 함수
def fetch_adgroup_stats(adgroup_ids):
    uri = "/stats"
    headers = get_header("GET", uri)

    # 필드 설정
    fields = [
        "clkCnt", "impCnt", "salesAmt", "ctr", "cpc",
        "ccnt", "crto", "convAmt", "ror", "cpConv", "viewCnt"
    ]

    # JSON 배열로 `fields`를 문자열로 변환
    fields_json = f"[{','.join(f'\"{field}\"' for field in fields)}]"

    # 쿼리 스트링 직접 생성
    params = {
        "datePreset": "yesterday",
        "fields": fields_json,  # JSON 배열로 전달
        "timeIncrement": "allDays",
        "ids": ",".join(adgroup_ids)  # 쉼표로 구분된 문자열
    }

    query_string = "&".join(f"{key}={urllib.parse.quote(value)}" for key, value in params.items())
    url = f"{BASE_URL}{uri}?{query_string}"

    print(f"[DEBUG] Final Encoded URL: {url}")

    try:
        # GET 요청
        response = requests.get(url, headers=headers)

        # 요청 정보 디버깅
        print(f"[DEBUG] Request Headers: {headers}")
        print(f"[DEBUG] Response Status: {response.status_code}")
        print(f"[DEBUG] Response Content: {response.text}")

        # 응답 처리
        response.raise_for_status()  # HTTP 에러 발생 시 예외 처리
        return response.json()
    except requests.exceptions.HTTPError as e:
        print(f"[ERROR] HTTP Error: {e.response.status_code}")
        print(f"[ERROR] Response Content: {e.response.text}")
        raise
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        raise

# 광고 소재 가져오기
def fetch_creatives_by_adgroup(adgroup_id):
    uri = "/ncc/ads"
    url = f"{BASE_URL}{uri}"
    headers = get_header("GET", uri)
    params = {"nccAdgroupId": adgroup_id}

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    creatives = response.json()

    # 활성화된 광고소재만 반환
    return [creative for creative in creatives if creative.get('status') == 'ELIGIBLE']

# 광고 소재 성과 가져오기
def fetch_creative_stats(creative_ids):
    uri = "/stats"
    headers = get_header("GET", uri)

    fields = [
        "clkCnt", "impCnt", "salesAmt", "ctr", "cpc",
        "ccnt", "crto", "convAmt", "ror", "cpConv", "viewCnt"
    ]
    fields_json = f"[{','.join(f'\"{field}\"' for field in fields)}]"

    params = {
        "datePreset": "yesterday",
        "fields": fields_json,
        "timeIncrement": "allDays",
        "ids": ",".join(creative_ids)
    }

    query_string = "&".join(f"{key}={urllib.parse.quote(value)}" for key, value in params.items())
    url = f"{BASE_URL}{uri}?{query_string}"

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Google Sheets 데이터 저장
def save_to_google_sheets(data):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    client = gspread.authorize(credentials)

    sheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1

    # 헤더 추가
    sheet.clear()
    sheet.append_row(["광고소재 ID", "클릭수", "노출수", "비용", "CTR", "CPC", "전환수", "전환율", "전환 금액", "ROAS", "전환당 비용", "조회수"])

    # 데이터를 한번에 추가
    rows = []
    for entry in data.get("data", []):
        rows.append([
            entry.get("id"),
            entry.get("clkCnt", 0),
            entry.get("impCnt", 0),
            entry.get("salesAmt", 0),
            entry.get("ctr", 0),
            entry.get("cpc", 0),
            entry.get("ccnt", 0),
            entry.get("crto", 0),
            entry.get("convAmt", 0),
            entry.get("ror", 0),
            entry.get("cpConv", 0),
            entry.get("viewCnt", 0)
        ])
    sheet.append_rows(rows, value_input_option="RAW")

# 메인 실행
if __name__ == "__main__":
    try:
        print("[INFO] Fetching active campaigns...")
        campaigns = fetch_campaigns()

        all_creative_stats = []

        for campaign in campaigns:
            print(f"[INFO] Processing campaign: {campaign['name']}")

            # 활성화된 광고그룹 가져오기
            adgroups = fetch_adgroups(campaign['nccCampaignId'])
            for adgroup in adgroups:
                print(f"[INFO] Processing adgroup: {adgroup['name']}")

                # 활성화된 광고소재 가져오기
                creatives = fetch_creatives_by_adgroup(adgroup['nccAdgroupId'])
                creative_ids = [creative['nccAdId'] for creative in creatives]

                if creative_ids:
                    # 광고소재 성과 가져오기
                    stats = fetch_creative_stats(creative_ids)
                    all_creative_stats.extend(stats.get("data", []))

        print("[INFO] Saving all stats to Google Sheets...")
        save_to_google_sheets({"data": all_creative_stats})

        print("[INFO] Completed successfully!")
    except Exception as e:
        print(f"[ERROR] {e}")