from facebook_business.api import FacebookAdsApi, FacebookAdsApiBatch
from facebook_business.adobjects.adaccount import AdAccount
from facebook_business.adobjects.campaign import Campaign
from facebook_business.adobjects.adset import AdSet
from facebook_business.adobjects.ad import Ad
from facebook_business.adobjects.adsinsights import AdsInsights
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta

# Facebook API 설정
ACCESS_TOKEN = 'EAAWZC8OFBBkUBOwVfRymBR5ucVX5m3jqAPftyRhC3ganIbIQIEfIfM9urcTnZC4zuXvBNgTVk6NNizIl5FhImnZACRcD7nx9ZBlcAEfZAZBZAslyTFlzK4vZADnzO8JsAKwWlD4JUDwXsMatVu4i8hcRsPtFqwM0lLh3Pq51NCGW86ZClAEhOsOfZAE82OIl5v0twQwAZDZD'
APP_ID = '1618416175744581'
APP_SECRET = 'dfe4c4294f787c3e0ce2f5187c5829a9' 
AD_ACCOUNT_ID = 'act_1080862949864051'

# Google Sheets API 설정
GOOGLE_SHEET_ID = '15q1Gqr3iyQM8DM6rjmPWyYO92r-1jo5ESV4DrLltbS4'
JSON_KEY_FILE = r'C:\Crolling\json\google_sheets_key.json'

# Facebook API 초기화
def initialize_api():
    FacebookAdsApi.init(app_id=APP_ID, app_secret=APP_SECRET, access_token=ACCESS_TOKEN)
    print("[INFO] Facebook API initialized successfully")

# 호출 횟수 관리
class APICallManager:
    def __init__(self, max_calls_per_second, cooldown_time):
        self.max_calls_per_second = 2  # 초당 호출 횟수 줄이기
        self.cooldown_time = cooldown_time
        self.call_count = 0
        self.start_time = time.time()

    def increment_and_check(self):
        self.call_count += 1
        elapsed_time = time.time() - self.start_time

        # 초당 호출 제한을 초과했는지 확인
        if elapsed_time < 1 and self.call_count >= self.max_calls_per_second:
            sleep_time = 1 - elapsed_time
            print(f"[INFO] API call limit reached ({self.call_count} calls/second). Sleeping for {sleep_time:.2f} seconds...")
            time.sleep(sleep_time)

        # 1초 단위로 호출 횟수 리셋
        if elapsed_time >= 1:
            self.call_count = 0
            self.start_time = time.time()

# 요청 실패 시 재시도 로직
def fetch_with_retry(fetch_function, max_retries=5, initial_delay=15, *args, **kwargs):
    delay = initial_delay
    for attempt in range(max_retries):
        try:
            return fetch_function(*args, **kwargs)
        except Exception as e:
            print(f"[ERROR] Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                print(f"[INFO] Retrying in {delay} seconds...")
                time.sleep(delay)
                delay = min(delay * 2, 60)  # 백오프 전략으로 대기 시간 두 배 증가, 최대 60초
            else:
                print("[ERROR] Maximum retries reached. Exiting...")
                raise

# 활성화된 캠페인 가져오기
def fetch_active_campaigns(ad_account_id, call_manager):
    ad_account = AdAccount(ad_account_id)
    campaigns = ad_account.get_campaigns(fields=[
        Campaign.Field.id,
        Campaign.Field.name,
        Campaign.Field.status
    ], params={"limit": 100})
    call_manager.increment_and_check()
    return [campaign for campaign in campaigns if campaign[Campaign.Field.status] == "ACTIVE"]

# 활성화된 광고 세트 가져오기
def fetch_active_adsets(campaign_id, call_manager):
    campaign = Campaign(campaign_id)
    adsets = campaign.get_ad_sets(fields=[
        AdSet.Field.id,
        AdSet.Field.name,
        AdSet.Field.daily_budget,
        AdSet.Field.status
    ], params={"limit": 100})
    call_manager.increment_and_check()
    return [adset for adset in adsets if adset[AdSet.Field.status] == "ACTIVE"]

# 활성화된 광고 가져오기
def fetch_active_ads(adset_id, call_manager):
    adset = AdSet(adset_id)
    ads = adset.get_ads(fields=[
        Ad.Field.id,
        Ad.Field.name,
        Ad.Field.status
    ], params={"limit": 100})
    call_manager.increment_and_check()
    return [ad for ad in ads if ad[Ad.Field.status] == "ACTIVE"]

# 광고 성과 데이터 가져오기 (전날 데이터)
def fetch_ad_insights(ad_id, call_manager):
    ad = Ad(ad_id)
    yesterday = datetime.now() - timedelta(days=1)
    date_preset = yesterday.strftime('%Y-%m-%d')
    insights = ad.get_insights(fields=[
        AdsInsights.Field.impressions,  # 노출
        AdsInsights.Field.spend,        # 지출 금액
        AdsInsights.Field.ctr,          # CTR
        AdsInsights.Field.purchase_roas  # 웹사이트 구매 ROAS
    ], params={"time_range": {"since": date_preset, "until": date_preset}})
    call_manager.increment_and_check()
    return insights

# Google Sheets에 데이터 저장
def save_to_google_sheets(data):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    client = gspread.authorize(credentials)

    sheet = client.open_by_key(GOOGLE_SHEET_ID)
    worksheet = sheet.worksheet("Sheet2")
    worksheet.clear()
    worksheet.append_row(["캠페인 이름", "광고 세트 이름", "광고 이름", "일일 예산", "ROAS", "CTR", "지출 금액", "노출"])

    rows = []
    for entry in data:
        rows.append([
            entry.get("campaign_name", ""),
            entry.get("adset_name", ""),
            entry.get("ad_name", ""),
            f"\\{int(entry.get('daily_budget', 0)):,}",
            f"{float(entry.get('roas', 0)):.2f}",
            f"{float(entry.get('ctr', 0)):.2f}%",
            f"\\{int(entry.get('spend', 0)):,}",
            f"{int(entry.get('impressions', 0)):,}"
        ])
    worksheet.append_rows(rows, value_input_option="RAW")

# 배치 요청을 위한 함수 추가
def batch_fetch_insights(ad_ids, call_manager):
    api = FacebookAdsApi.get_default_api()
    batch = FacebookAdsApiBatch(api)
    insights_map = {}
    
    yesterday = datetime.now() - timedelta(days=1)
    date_preset = yesterday.strftime('%Y-%m-%d')
    
    for ad_id in ad_ids:
        ad = Ad(ad_id)
        insights = ad.get_insights(
            fields=[
                AdsInsights.Field.impressions,
                AdsInsights.Field.spend,
                AdsInsights.Field.ctr,
                AdsInsights.Field.purchase_roas
            ],
            params={"time_range": {"since": date_preset, "until": date_preset}}
        )
        batch.add_request(insights)
    
    batch_responses = batch.execute()
    return batch_responses

# 메인 실행 부분 수정
def process_ads_in_batches(ads, batch_size=50):
    ad_ids = [ad['id'] for ad in ads]
    results = []
    
    for i in range(0, len(ad_ids), batch_size):
        batch_ads = ad_ids[i:i + batch_size]
        batch_results = batch_fetch_insights(batch_ads, api_call_manager)
        results.extend(batch_results)
    
    return results

# 메인 실행
if __name__ == "__main__":
    try:
        initialize_api()
        print("[INFO] Fetching active campaigns...")

        # API 호출 관리 객체 생성
        api_call_manager = APICallManager(max_calls_per_second=2, cooldown_time=120)

        # 활성화된 캠페인 가져오기
        campaigns = fetch_with_retry(fetch_active_campaigns, ad_account_id=AD_ACCOUNT_ID, call_manager=api_call_manager)

        all_data = []
        for campaign in campaigns:
            campaign_name = campaign[Campaign.Field.name]
            print(f"[INFO] Processing campaign: {campaign_name}")

            # 활성화된 광고 세트 가져오기
            adsets = fetch_with_retry(fetch_active_adsets, campaign_id=campaign[Campaign.Field.id], call_manager=api_call_manager)
            for adset in adsets:
                adset_name = adset[AdSet.Field.name]
                daily_budget = adset.get(AdSet.Field.daily_budget, 0)

                # 활성화된 광고 가져오기
                ads = fetch_with_retry(fetch_active_ads, adset_id=adset[AdSet.Field.id], call_manager=api_call_manager)
                for ad in ads:
                    ad_name = ad[Ad.Field.name]

                    # 광고 성과 데이터 가져오기
                    insights = fetch_with_retry(fetch_ad_insights, ad_id=ad[Ad.Field.id], call_manager=api_call_manager)
                    for insight in insights:
                        roas = insight.get("purchase_roas", [{}])[0].get("value", 0)
                        ctr = insight.get("ctr", 0)
                        spend = insight.get("spend", 0)
                        impressions = insight.get("impressions", 0)

                        print(f"[INFO] Processing ad: {ad_name} (ROAS: {roas}, CTR: {ctr}, Spend: {spend}, Impressions: {impressions})")
                        all_data.append({
                            "campaign_name": campaign_name,
                            "adset_name": adset_name,
                            "ad_name": ad_name,
                            "daily_budget": daily_budget,
                            "roas": roas,
                            "ctr": ctr,
                            "spend": spend,
                            "impressions": impressions
                        })

        print("[INFO] Saving data to Google Sheets...")
        save_to_google_sheets(all_data)
        print("[INFO] Completed successfully!")

    except Exception as e:
        print(f"[ERROR] {e}")
