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

    return {
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Timestamp': timestamp,
        'X-API-KEY': NAVER_CLIENT_KEY,
        'X-Customer': NAVER_CUSTOMER_ID,
        'X-Signature': signature
    }

# 활성화된 캠페인 가져오기
def fetch_campaigns():
    uri = "/ncc/campaigns"
    url = f"{BASE_URL}{uri}"
    headers = get_header("GET", uri)

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    campaigns = response.json()

    return [campaign for campaign in campaigns if campaign.get('status') == 'ELIGIBLE']

# 광고그룹 가져오기
def fetch_adgroups(campaign_id):
    uri = "/ncc/adgroups"
    url = f"{BASE_URL}{uri}"
    headers = get_header("GET", uri)
    params = {"nccCampaignId": campaign_id}

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    adgroups = response.json()

    return [adgroup for adgroup in adgroups if adgroup.get('status') == 'ELIGIBLE']

# 광고소재 가져오기
def fetch_creatives_by_adgroup(adgroup_id):
    uri = "/ncc/ads"
    url = f"{BASE_URL}{uri}"
    headers = get_header("GET", uri)
    params = {"nccAdgroupId": adgroup_id}

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    creatives = response.json()

    result = []
    for creative in creatives:
        if creative.get("status") == "ELIGIBLE":
            # referenceData에서 productName 또는 productTitle 가져오기
            reference_data = creative.get("referenceData", {})
            product_name = reference_data.get("productName", "Unknown Product")  # 기본 상품명

            result.append({
                "nccAdId": creative.get("nccAdId"),
                "adgroupName": creative.get("nccAdgroupId"),
                "productName": product_name,  # 광고소재 기본 상품명
            })
    return result

# 광고소재 성과 가져오기
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

# 중복 제거 함수 (광고그룹 변경 시 캠페인 이름 표시)
def remove_duplicates_with_group_check(data):
    """
    연속된 중복 캠페인 이름을 제거하지만, 광고그룹이 변경되면 다시 캠페인 이름을 표시합니다.
    :param data: 성과 데이터 리스트
    :return: 중복 제거된 데이터 리스트
    """
    previous_campaign = None
    previous_adgroup = None

    for row in data:
        # 광고그룹 변경 시 캠페인 이름 표시
        if row["adgroupName"] != previous_adgroup:
            # 광고그룹이 달라지면 캠페인 이름 유지
            previous_adgroup = row["adgroupName"]
        else:
            # 광고그룹이 같으면 캠페인 이름 비움
            row["campaignName"] = ""

        # 광고그룹 이름 중복 제거
        if row["adgroupName"] == previous_adgroup:
            if row["campaignName"] == "":
                row["adgroupName"] = ""
        else:
            previous_adgroup = row["adgroupName"]

        # 캠페인 이름 처리
        if row["campaignName"] == previous_campaign:
            row["campaignName"] = ""
        else:
            previous_campaign = row["campaignName"]

    return data


# Google Sheets 데이터 저장
def save_to_google_sheets(data):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    client = gspread.authorize(credentials)

    sheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1

    # 헤더 추가
    sheet.clear()
    sheet.append_row([
        "캠페인 이름", "광고그룹 이름", "광고소재 이름", "노출수", "클릭수", "CTR", "CPC",
        "총비용", "전환수", "전환율", "전환매출액", "ROAS", "전환당비용"
    ])

    # 중복 제거
    cleaned_data = remove_duplicates_with_group_check(data.get("data", []))

    # 데이터를 한번에 추가
    rows = []
    for entry in cleaned_data:
        rows.append([
            entry.get("campaignName", ""),  # 캠페인 이름
            entry.get("adgroupName", ""),  # 광고그룹 이름
            entry.get("productName", ""),  # 광고소재 이름
            entry.get("impCnt", 0),  # 노출수
            entry.get("clkCnt", 0),  # 클릭수
            entry.get("ctr", 0),  # CTR
            entry.get("cpc", 0),  # CPC
            entry.get("salesAmt", 0),  # 총비용
            entry.get("ccnt", 0),  # 전환수
            entry.get("crto", 0),  # 전환율
            entry.get("convAmt", 0),  # 전환매출액
            entry.get("ror", 0),  # ROAS
            entry.get("cpConv", 0)  # 전환당비용
        ])
    sheet.append_rows(rows, value_input_option="RAW")

# 메인 실행
if __name__ == "__main__":
    try:
        print("[INFO] Fetching active campaigns...")
        # 활성화된 캠페인 가져오기
        campaigns = fetch_campaigns()

        # 모든 광고소재 통계 데이터를 저장할 리스트
        all_creative_stats = []

        # 캠페인을 순회하며 데이터를 가져옴
        for campaign in campaigns:
            print(f"[INFO] Processing campaign: {campaign['name']}")

            # 광고그룹 가져오기
            adgroups = fetch_adgroups(campaign['nccCampaignId'])
            for adgroup in adgroups:
                print(f"[INFO] Processing adgroup: {adgroup['name']}")

                # 광고소재 가져오기
                creatives = fetch_creatives_by_adgroup(adgroup['nccAdgroupId'])
                creative_ids = [creative["nccAdId"] for creative in creatives]

                if creative_ids:
                    # 광고소재 성과 가져오기 (with retry)
                    creative_stats = fetch_creative_stats(creative_ids)

                    # 성과 데이터를 정리하여 저장
                    for stat in creative_stats.get("data", []):
                        stat["campaignName"] = campaign["name"]  # 캠페인 이름 추가
                        stat["adgroupName"] = adgroup["name"]  # 광고그룹 이름 추가
                        stat["productName"] = next(
                            (creative["productName"] for creative in creatives if creative["nccAdId"] == stat["id"]),
                            "Unknown Product"  # 기본 상품명을 찾지 못한 경우
                        )
                        all_creative_stats.append(stat)

        # 모든 데이터를 Google Sheets에 저장
        print("[INFO] Saving all stats to Google Sheets...")
        save_to_google_sheets({"data": all_creative_stats})

        print("[INFO] Completed successfully!")

    except Exception as e:
        print(f"[ERROR] {e}")