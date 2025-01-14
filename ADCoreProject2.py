from facebook_business.api import FacebookAdsApi, FacebookAdsApiBatch
from facebook_business.adobjects.adaccount import AdAccount
from facebook_business.adobjects.campaign import Campaign
from facebook_business.adobjects.adset import AdSet
from facebook_business.adobjects.adsinsights import AdsInsights
from facebook_business.adobjects.ad import Ad
from facebook_business.exceptions import FacebookRequestError
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
from gspread_formatting import CellFormat, format_cell_range, Color, TextFormat
import random

# Facebook API 설정
ACCESS_TOKEN = 'EAAWZC8OFBBkUBOwVfRymBR5ucVX5m3jqAPftyRhC3ganIbIQIEfIfM9urcTnZC4zuXvBNgTVk6NNizIl5FhImnZACRcD7nx9ZBlcAEfZAZBZAslyTFlzK4vZADnzO8JsAKwWlD4JUDwXsMatVu4i8hcRsPtFqwM0lLh3Pq51NCGW86ZClAEhOsOfZAE82OIl5v0twQwAZDZD'
APP_ID = '1618416175744581'
APP_SECRET = 'dfe4c4294f787c3e0ce2f5187c5829a9'
AD_ACCOUNT_ID = 'act_1080862949864051'

# Google Sheets API 설정
GOOGLE_SHEET_ID = '15q1Gqr3iyQM8DM6rjmPWyYO92r-1jo5ESV4DrLltbS4'
JSON_KEY_FILE = r'C:\selenium\json\google_sheets_key.json'

# Facebook API 초기화
def initialize_api():
    FacebookAdsApi.init(app_id=APP_ID, app_secret=APP_SECRET, access_token=ACCESS_TOKEN)
    print("[INFO] Facebook API initialized successfully")

# 호출 횟수 관리
class APICallManager:
    def __init__(self, max_calls_per_second=1, cooldown_time=2):
        self.max_calls_per_second = max_calls_per_second  # 초당 호출 횟수 제한
        self.cooldown_time = cooldown_time  # 쿨다운 시간
        self.call_count = 0
        self.start_time = time.time()

    def increment_and_check(self):
        self.call_count += 1
        elapsed_time = time.time() - self.start_time

        # 호출 속도 제한: 5~10초 사이로 설정
        delay_time = random.uniform(5, 10)

        if elapsed_time < delay_time:  # 호출 간 대기 시간이 부족한 경우
            sleep_time = delay_time - elapsed_time
            print(f"[INFO] API call limit reached. Sleeping for {sleep_time:.2f} seconds...")
            time.sleep(sleep_time)

        self.last_call_time = time.time()  # 마지막 호출 시간 갱신

# 요청 실패 시 재시도 로직
def fetch_with_retry(fetch_function, max_retries=5, initial_delay=60, *args, **kwargs):
    delay = initial_delay
    for attempt in range(max_retries):
        try:
            return fetch_function(*args, **kwargs)
        except Exception as e:
            print(f"[ERROR] Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                print(f"[INFO] Retrying in {delay} seconds...")
                time.sleep(delay)
                delay = min(delay * 2, 300)  # 최대 5분 대기
            else:
                print("[ERROR] Maximum retries reached. Exiting...")
                raise
            
# 배치 요청으로 광고 성과 데이터 가져오기
def batch_fetch_insights(ad_ids, call_manager, batch_size=50):
    api = FacebookAdsApi.get_default_api()
    all_insights = []
    yesterday = datetime.now() - timedelta(days=1)
    date_preset = {"since": yesterday.strftime('%Y-%m-%d'), "until": yesterday.strftime('%Y-%m-%d')}

    for i in range(0, len(ad_ids), batch_size):
        batch = FacebookAdsApiBatch(api)
        batch_ads = ad_ids[i:i + batch_size]

        for ad_id in batch_ads:
            # 올바른 요청 생성
            batch.add_request(
                Ad(ad_id).get_insights(
                    fields=[
                        AdsInsights.Field.impressions,
                        AdsInsights.Field.spend,
                        AdsInsights.Field.ctr,
                        AdsInsights.Field.purchase_roas,
                    ],
                    params={"time_range": date_preset}
                ),
                success=lambda response: all_insights.extend(response),
                failure=lambda error: print(f"[ERROR] Failed for ad {ad_id}: {error}")
            )

        try:
            batch.execute()
            call_manager.increment_and_check()
        except FacebookRequestError as e:
            print(f"[ERROR] Batch execution failed: {str(e)}")
            continue

    return all_insights


# 활성화된 캠페인 가져오기
def fetch_active_campaigns(ad_account_id, call_manager):
    ad_account = AdAccount(ad_account_id)
    campaigns = list(ad_account.get_campaigns(fields=[
        Campaign.Field.id,
        Campaign.Field.name,
        Campaign.Field.status
    ], params={"limit": 100}))
    call_manager.increment_and_check()
    return [campaign for campaign in campaigns if campaign[Campaign.Field.status] == "ACTIVE"]

# 활성화된 광고 세트 가져오기
def fetch_active_adsets(campaign_id, call_manager):
    try:
        campaign = Campaign(campaign_id)
        adsets = list(campaign.get_ad_sets(fields=[
            AdSet.Field.id,
            AdSet.Field.name,
            AdSet.Field.daily_budget,
            AdSet.Field.status
        ], params={"limit": 100}))
        call_manager.increment_and_check()
        return [adset for adset in adsets if adset[AdSet.Field.status] == "ACTIVE"]
    except Exception as e:
        print(f"[ERROR] Failed to fetch adsets for campaign {campaign_id}: {e}")
        raise

# 활성화된 광고 가져오기
def fetch_active_ads(adset_id, call_manager):
    try:
        adset = AdSet(adset_id)
        ads = list(adset.get_ads(fields=[
            Ad.Field.id,
            Ad.Field.name,
            Ad.Field.status
        ], params={"limit": 100}))
        call_manager.increment_and_check()
        return [ad for ad in ads if ad[Ad.Field.status] == "ACTIVE"]
    except Exception as e:
        print(f"[ERROR] Failed to fetch ads for adset {adset_id}: {e}")
        raise


# 광고 성과 데이터 가져오기 (전날 데이터)
# def fetch_ad_insights(ad_id, call_manager):
#     try:
#         ad = Ad(ad_id)
#         yesterday = datetime.now() - timedelta(days=1)
#         date_preset = yesterday.strftime('%Y-%m-%d')
#         insights = ad.get_insights(fields=[
#             AdsInsights.Field.impressions,
#             AdsInsights.Field.spend,
#             AdsInsights.Field.ctr,
#             AdsInsights.Field.purchase_roas
#         ], params={"time_range": {"since": date_preset, "until": date_preset}})
#         call_manager.increment_and_check()
#         return insights
#     except Exception as e:
#         print(f"[ERROR] Failed to fetch insights for ad {ad_id}: {e}")
#         raise

# Google Sheets에 데이터 저장
def save_to_google_sheets(data):
    # Google Sheets API 인증 설정
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    client = gspread.authorize(credentials)

    # Google Sheet 열기 및 초기화
    sheet = client.open_by_key(GOOGLE_SHEET_ID)
    worksheet = sheet.worksheet("Sheet2")
    worksheet.clear()

    # 기본 배경색 설정
    white_cell_format = CellFormat(backgroundColor=Color(1, 1, 1))
    format_cell_range(worksheet, "A1:H1000", white_cell_format)

    # 제목 행 추가
    yesterday = (datetime.now() - timedelta(days=1)).strftime('%y.%m.%d')
    worksheet.append_row([f"캠페인이름({yesterday})", "광고 세트 이름", "광고 이름", "일일 예산", "ROAS", "CTR", "지출 금액", "노출"])

    # 데이터 처리에 필요한 변수 초기화
    rows = []  # 최종 데이터를 저장할 리스트
    adset_ranges = []  # 광고 세트 이름 총합 행 범위 저장
    campaign_ranges = []  # 캠페인 이름 총합 행 범위 저장
    overall_range = None  # 전체 캠페인 총합 행 저장

    # 각 총합 계산을 위한 변수
    current_adset_name = ""
    current_campaign_name = ""
    adset_total = {"daily_budget": 0, "roas": 0, "ctr": 0, "spend": 0, "impressions": 0, "count": 0}
    campaign_total = {"daily_budget": 0, "roas": 0, "ctr": 0, "spend": 0, "impressions": 0, "count": 0}
    overall_total = {"daily_budget": 0, "roas": 0, "ctr": 0, "spend": 0, "impressions": 0, "count": 0}

    for entry in data:
        # 광고 세트 이름 변경 시, 광고 세트 총합 행 추가
        if current_adset_name and current_adset_name != entry["adset_name"]:
            rows.append([  # 광고 세트 이름 총합 행 추가
                "", current_adset_name + " 총합", "",
                f"₩{adset_total['daily_budget']:,}" if adset_total['daily_budget'] else "",
                f"{(adset_total['roas'] / adset_total['count']):.2f}" if adset_total['count'] > 0 else "",
                f"{(adset_total['ctr'] / adset_total['count']):.2f}%" if adset_total['count'] > 0 else "",
                f"₩{adset_total['spend']:,}" if adset_total['spend'] else "",
                f"{adset_total['impressions']:,}" if adset_total['impressions'] else ""
            ])
            adset_ranges.append(len(rows))  # 광고 세트 총합 행의 위치 저장
            adset_total = {"daily_budget": 0, "roas": 0, "ctr": 0, "spend": 0, "impressions": 0, "count": 0}

        # 캠페인 이름 변경 시, 캠페인 이름 총합 행 추가
        if current_campaign_name and current_campaign_name != entry["campaign_name"]:
            rows.append([  # 캠페인 이름 총합 행 추가
                current_campaign_name + " 총합", "", "",
                f"₩{campaign_total['daily_budget']:,}" if campaign_total['daily_budget'] else "",
                f"{(campaign_total['roas'] / campaign_total['count']):.2f}" if campaign_total['count'] > 0 else "",
                f"{(campaign_total['ctr'] / campaign_total['count']):.2f}%" if campaign_total['count'] > 0 else "",
                f"₩{campaign_total['spend']:,}" if campaign_total['spend'] else "",
                f"{campaign_total['impressions']:,}" if campaign_total['impressions'] else ""
            ])
            campaign_ranges.append(len(rows))  # 캠페인 이름 총합 행의 위치 저장
            campaign_total = {"daily_budget": 0, "roas": 0, "ctr": 0, "spend": 0, "impressions": 0, "count": 0}

        # 현재 광고 세트 이름 및 캠페인 이름 갱신
        current_adset_name = entry["adset_name"]
        current_campaign_name = entry["campaign_name"]

        # 광고 세트 및 캠페인, 전체 합계 업데이트
        adset_total["daily_budget"] += int(entry.get("daily_budget", 0))
        adset_total["roas"] += float(entry.get("roas", 0))
        adset_total["ctr"] += float(entry.get("ctr", 0))
        adset_total["spend"] += int(entry.get("spend", 0))
        adset_total["impressions"] += int(entry.get("impressions", 0))
        adset_total["count"] += 1

        campaign_total["daily_budget"] += int(entry.get("daily_budget", 0))
        campaign_total["roas"] += float(entry.get("roas", 0))
        campaign_total["ctr"] += float(entry.get("ctr", 0))
        campaign_total["spend"] += int(entry.get("spend", 0))
        campaign_total["impressions"] += int(entry.get("impressions", 0))
        campaign_total["count"] += 1

        overall_total["daily_budget"] += int(entry.get("daily_budget", 0))
        overall_total["roas"] += float(entry.get("roas", 0))
        overall_total["ctr"] += float(entry.get("ctr", 0))
        overall_total["spend"] += int(entry.get("spend", 0))
        overall_total["impressions"] += int(entry.get("impressions", 0))
        overall_total["count"] += 1

        # 데이터 행 추가
        rows.append([entry.get("campaign_name", ""), entry.get("adset_name", ""), entry.get("ad_name", ""),
                    f"₩{int(entry.get('daily_budget', 0)):,}" if entry.get('daily_budget') else "",
                    f"{float(entry.get('roas', 0)):.2f}" if entry.get("roas") else "",
                    f"{float(entry.get('ctr', 0)):.2f}%" if entry.get("ctr") else "",
                    f"₩{int(entry.get('spend', 0)):,}" if entry.get("spend") else "",
                    f"{int(entry.get('impressions', 0)):,}" if entry.get("impressions") else ""])

    # 마지막 광고 세트 및 캠페인 총합 추가
    if current_adset_name:
        rows.append([  # 광고 세트 이름 총합
            "", current_adset_name + " 총합", "",
            f"₩{adset_total['daily_budget']:,}" if adset_total['daily_budget'] else "",
            f"{(adset_total['roas'] / adset_total['count']):.2f}" if adset_total['count'] > 0 else "",
            f"{(adset_total['ctr'] / adset_total['count']):.2f}%" if adset_total['count'] > 0 else "",
            f"₩{adset_total['spend']:,}" if adset_total['spend'] else "",
            f"{adset_total['impressions']:,}" if adset_total['impressions'] else ""
        ])
        adset_ranges.append(len(rows))  # 광고 세트 총합 행의 위치 저장

    if current_campaign_name:
        rows.append([  # 캠페인 이름 총합
            current_campaign_name + " 총합", "", "",
            f"₩{campaign_total['daily_budget']:,}" if campaign_total['daily_budget'] else "",
            f"{(campaign_total['roas'] / campaign_total['count']):.2f}" if campaign_total['count'] > 0 else "",
            f"{(campaign_total['ctr'] / campaign_total['count']):.2f}%" if campaign_total['count'] > 0 else "",
            f"₩{campaign_total['spend']:,}" if campaign_total['spend'] else "",
            f"{campaign_total['impressions']:,}" if campaign_total['impressions'] else ""
        ])
        campaign_ranges.append(len(rows))  # 캠페인 이름 총합 행의 위치 저장

    # 전체 캠페인 총합 행 추가
    rows.append([
        "전체 캠페인 총합", "", "",
        f"₩{overall_total['daily_budget']:,}" if overall_total['daily_budget'] else "",
        f"{(overall_total['roas'] / overall_total['count']):.2f}" if overall_total['count'] > 0 else "",
        f"{(overall_total['ctr'] / overall_total['count']):.2f}%" if overall_total['count'] > 0 else "",
        f"₩{overall_total['spend']:,}" if overall_total['spend'] else "",
        f"{overall_total['impressions']:,}" if overall_total['impressions'] else ""
    ])
    overall_range = len(rows)  # 전체 캠페인 총합 행의 위치 저장

    # 모든 데이터를 Google Sheets에 추가
    worksheet.append_rows(rows, value_input_option="RAW")

    # 강조 색상 설정
    campaign_highlight_format = CellFormat(backgroundColor=Color(1, 1, 0))  # 노란색
    adset_highlight_format = CellFormat(backgroundColor=Color(0, 1, 0))  # 초록색
    overall_highlight_format = CellFormat(backgroundColor=Color(0.8, 0.5, 0.8))  # 보라색

    # 광고 세트 이름 총합 (B, D-H 열만 칠함)
    for row in adset_ranges:
        format_cell_range(worksheet, f"B{row+1}", adset_highlight_format)
        format_cell_range(worksheet, f"D{row+1}:H{row+1}", adset_highlight_format)

    # 캠페인 이름 총합 (A, D-H 열만 칠함)
    for row in campaign_ranges:
        format_cell_range(worksheet, f"A{row+1}", campaign_highlight_format)
        format_cell_range(worksheet, f"D{row+1}:H{row+1}", campaign_highlight_format)

    # 전체 캠페인 총합 (A, D-H 열만 칠함)
    if overall_range:
        format_cell_range(worksheet, f"A{overall_range+1}", overall_highlight_format)
        format_cell_range(worksheet, f"D{overall_range+1}:H{overall_range+1}", overall_highlight_format)

if __name__ == "__main__":
    try:
        initialize_api()
        print("[INFO] Fetching active campaigns...")

        api_call_manager = APICallManager(max_calls_per_second=1, cooldown_time=300)
        campaigns = fetch_with_retry(fetch_active_campaigns, ad_account_id=AD_ACCOUNT_ID, call_manager=api_call_manager)

        all_data = []
        for campaign in campaigns:
            try:
                campaign_name = campaign[Campaign.Field.name]
                print(f"[INFO] Processing campaign: {campaign_name}")

                adsets = fetch_with_retry(fetch_active_adsets, campaign_id=campaign[Campaign.Field.id], call_manager=api_call_manager)
                for adset in adsets:
                    adset_name = adset[AdSet.Field.name]
                    daily_budget = adset.get(AdSet.Field.daily_budget, 0)

                    ads = fetch_with_retry(fetch_active_ads, adset_id=adset[AdSet.Field.id], call_manager=api_call_manager)
                    ad_ids = [ad[Ad.Field.id] for ad in ads]

                    # 배치 요청으로 광고 성과 데이터 가져오기
                    insights = batch_fetch_insights(ad_ids, api_call_manager)
                    for insight in insights:
                        all_data.append({
                            "campaign_name": campaign_name,
                            "adset_name": adset_name,
                            "ad_name": insight.get("ad_name", ""),
                            "daily_budget": daily_budget,
                            "roas": insight.get("purchase_roas", [{}])[0].get("value", 0),
                            "ctr": insight.get("ctr", 0),
                            "spend": insight.get("spend", 0),
                            "impressions": insight.get("impressions", 0)
                        })
            except Exception as e:
                print(f"[ERROR] Failed to process campaign: {campaign_name}. Error: {e}")
                continue

        print("[INFO] Saving data to Google Sheets...")
        save_to_google_sheets(all_data)
        print("[INFO] Completed successfully!")

    except Exception as e:
        print(f"[ERROR] {e}") 