from facebook_business.api import FacebookAdsApi
from facebook_business.adobjects.adaccount import AdAccount
from facebook_business.adobjects.campaign import Campaign
from facebook_business.adobjects.adset import AdSet
from facebook_business.adobjects.ad import Ad
from facebook_business.adobjects.adsinsights import AdsInsights
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials

# Facebook API 설정
ACCESS_TOKEN = 'EAAWZC8OFBBkUBOwVfRymBR5ucVX5m3jqAPftyRhC3ganIbIQIEfIfM9urcTnZC4zuXvBNgTVk6NNizIl5FhImnZACRcD7nx9ZBlcAEfZAZBZAslyTFlzK4vZADnzO8JsAKwWlD4JUDwXsMatVu4i8hcRsPtFqwM0lLh3Pq51NCGW86ZClAEhOsOfZAE82OIl5v0twQwAZDZD'
APP_ID = '1618416175744581'
APP_SECRET = 'dfe4c4294f787c3e0ce2f5187c5829a9'
AD_ACCOUNT_ID = 'act_1080862949864051'  # 광고 계정 ID

# Google Sheets API 설정
GOOGLE_SHEET_ID = '15q1Gqr3iyQM8DM6rjmPWyYO92r-1jo5ESV4DrLltbS4'
JSON_KEY_FILE = r'C:\Crolling\json\google_sheets_key.json'

# Facebook Ads API 초기화
def initialize_api():
    try:
        FacebookAdsApi.init(app_id=APP_ID, app_secret=APP_SECRET, access_token=ACCESS_TOKEN)
        print("[INFO] Facebook API initialized successfully")
    except Exception as e:
        print(f"[ERROR] Failed to initialize Facebook API: {e}")
        raise

# 요청 실패 시 재시도 로직
def fetch_with_retry(fetch_function, max_retries=5, delay=10, *args, **kwargs):
    for attempt in range(max_retries):
        try:
            return fetch_function(*args, **kwargs)
        except Exception as e:
            print(f"[ERROR] Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                print(f"[INFO] Retrying in {delay} seconds...")
                time.sleep(delay)
            else:
                print("[ERROR] Maximum retries reached. Exiting...")
                raise

# 캠페인 가져오기
def fetch_campaigns(ad_account_id):
    ad_account = AdAccount(ad_account_id)
    return ad_account.get_campaigns(fields=[Campaign.Field.id, Campaign.Field.name, Campaign.Field.status])

# 광고 세트 가져오기
def fetch_adsets(campaign_id):
    campaign = Campaign(campaign_id)
    adsets_cursor = campaign.get_ad_sets(fields=[AdSet.Field.id, AdSet.Field.name, AdSet.Field.status])
    for adset in adsets_cursor:
        yield adset

# 광고 가져오기
def fetch_ads(adset_id):
    adset = AdSet(adset_id)
    ads_cursor = adset.get_ads(fields=[Ad.Field.id, Ad.Field.name, Ad.Field.status])
    for ad in ads_cursor:
        yield ad

# 광고 성과 가져오기
def fetch_insights(ad_id):
    ad = Ad(ad_id)
    return ad.get_insights(fields=[AdsInsights.Field.impressions, AdsInsights.Field.clicks, AdsInsights.Field.spend, AdsInsights.Field.ctr])

# Google Sheets에 데이터 저장
def save_to_google_sheets(data):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    client = gspread.authorize(credentials)

    sheet = client.open_by_key(GOOGLE_SHEET_ID)
    worksheet = sheet.worksheet("Sheet2")
    worksheet.clear()
    worksheet.append_row(["캠페인 이름", "광고 세트 이름", "광고 이름", "노출수", "클릭수", "CTR", "광고 비용"])

    rows = []
    for entry in data:
        rows.append([
            entry.get("campaign_name", ""),
            entry.get("adset_name", ""),
            entry.get("ad_name", ""),
            entry.get("impressions", 0),
            entry.get("clicks", 0),
            entry.get("ctr", 0),
            entry.get("spend", 0),
        ])
    worksheet.append_rows(rows, value_input_option="RAW")

# 메인 실행
if __name__ == "__main__":
    try:
        initialize_api()

        print("[INFO] Fetching campaigns...")
        campaigns = fetch_with_retry(fetch_campaigns, ad_account_id=AD_ACCOUNT_ID)

        all_data = []
        for campaign in campaigns:
            campaign_name = campaign[Campaign.Field.name]
            print(f"[INFO] Processing campaign: {campaign_name}")

            for adset in fetch_adsets(campaign[Campaign.Field.id]):
                adset_name = adset[AdSet.Field.name]
                print(f"[INFO] Processing ad set: {adset_name}")
                time.sleep(10)  # 광고 세트 간 대기 시간

                for ad in fetch_ads(adset[AdSet.Field.id]):
                    ad_name = ad[Ad.Field.name]
                    print(f"[INFO] Processing ad: {ad_name}")
                    time.sleep(5)  # 광고 간 대기 시간

                    insights = fetch_with_retry(fetch_insights, ad_id=ad[Ad.Field.id])
                    for insight in insights:
                        all_data.append({
                            "campaign_name": campaign_name,
                            "adset_name": adset_name,
                            "ad_name": ad_name,
                            "impressions": insight.get("impressions", 0),
                            "clicks": insight.get("clicks", 0),
                            "ctr": insight.get("ctr", 0),
                            "spend": insight.get("spend", 0),
                        })

        print("[INFO] Saving data to Google Sheets...")
        save_to_google_sheets(all_data)
        print("[INFO] Completed successfully!")

    except Exception as e:
        print(f"[ERROR] {e}")