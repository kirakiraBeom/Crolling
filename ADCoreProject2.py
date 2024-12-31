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
from gspread_formatting import CellFormat, format_cell_range, Color
import json
import os
import sys

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
        self.max_calls_per_second = 1  # 초당 호출 횟수 줄이기
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
def fetch_with_retry(fetch_function, max_retries=5, initial_delay=30, *args, **kwargs):
    delay = initial_delay
    for attempt in range(max_retries):
        try:
            return fetch_function(*args, **kwargs)
        except Exception as e:
            print(f"[ERROR] Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                print(f"[INFO] Retrying in {delay} seconds...")
                time.sleep(delay)
                delay = min(delay * 2, 120)  # 백오프 전략으로 대기 시간 두 배 증가, 최대 120초
            else:
                print("[ERROR] Maximum retries reached. Exiting...")
                raise

# 활성화된 캠페인 가져오기
def fetch_active_campaigns(ad_account_id, call_manager):
    try:
        ad_account = AdAccount(ad_account_id)
        campaigns = ad_account.get_campaigns(fields=[
            Campaign.Field.id,
            Campaign.Field.name,
            Campaign.Field.status
        ], params={"limit": 100})
        call_manager.increment_and_check()
        return [campaign for campaign in campaigns if campaign[Campaign.Field.status] == "ACTIVE"]
    except Exception as e:
        print(f"[ERROR] Campaign fetch failed: {e}")
        raise

# 활성화된 광고 세트 가져오기
def fetch_active_adsets(campaign_id, call_manager):
    try:
        campaign = Campaign(campaign_id)
        adsets = campaign.get_ad_sets(fields=[
            AdSet.Field.id,
            AdSet.Field.name,
            AdSet.Field.daily_budget,
            AdSet.Field.status,
            AdSet.Field.learning_stage_info
        ], params={"limit": 100})
        call_manager.increment_and_check()
        return [
            adset for adset in adsets
            if adset[AdSet.Field.status] == "ACTIVE" or adset.get("learning_stage_info", {}).get("status", "") == "LEARNING"
        ]
    except Exception as e:
        print(f"[ERROR] AdSet fetch failed: {e}")
        raise

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
    try:
        ad = Ad(ad_id)
        yesterday = datetime.now() - timedelta(days=1)
        date_preset = yesterday.strftime('%Y-%m-%d')
        insights = ad.get_insights(fields=[
            AdsInsights.Field.impressions,
            AdsInsights.Field.spend,
            AdsInsights.Field.ctr,
            AdsInsights.Field.purchase_roas
        ], params={"time_range": {"since": date_preset, "until": date_preset}})
        call_manager.increment_and_check()
        return insights
    except Exception as e:
        print(f"[ERROR] Ad insights fetch failed: {e}")
        raise

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
    current_campaign_name = ""
    current_adset_name = ""

    campaign_total = {
        "daily_budget": 0,
        "roas": 0,
        "ctr": 0,
        "spend": 0,
        "impressions": 0,
        "count": 0
    }

    adset_total = {
        "daily_budget": 0,
        "roas": 0,
        "ctr": 0,
        "spend": 0,
        "impressions": 0,
        "count": 0
    }

    for entry in data:
        if current_campaign_name != entry["campaign_name"]:
            if current_campaign_name:
                rows.append([current_campaign_name + " 총합", "", "", f"₩{campaign_total['daily_budget']:,}", f"{(campaign_total['roas'] / campaign_total['count']):.2f}", f"{(campaign_total['ctr'] / campaign_total['count']):.2f}%", f"₩{campaign_total['spend']:,}", f"{campaign_total['impressions']:,}"])
            current_campaign_name = entry["campaign_name"]
            campaign_total = {"daily_budget": 0, "roas": 0, "ctr": 0, "spend": 0, "impressions": 0, "count": 0}
            rows.append([current_campaign_name, "", "", "", "", "", "", ""])

        if current_adset_name != entry["adset_name"]:
            if current_adset_name:
                rows.append(["", current_adset_name + " 총합", "", f"₩{adset_total['daily_budget']:,}", f"{(adset_total['roas'] / adset_total['count']):.2f}", f"{(adset_total['ctr'] / adset_total['count']):.2f}%", f"₩{adset_total['spend']:,}", f"{adset_total['impressions']:,}"])
            current_adset_name = entry["adset_name"]
            adset_total = {"daily_budget": 0, "roas": 0, "ctr": 0, "spend": 0, "impressions": 0, "count": 0}
            rows.append(["", current_adset_name, "", "", "", "", "", ""])

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

        rows.append([
            "",
            "",
            entry.get("ad_name", ""),
            f"₩{int(entry.get('daily_budget', 0)):,}" if entry.get('daily_budget') else "-",
            f"{float(entry.get('roas', 0)):.2f}",
            f"{float(entry.get('ctr', 0)):.2f}%",
            f"₩{int(entry.get('spend', 0)):,}",
            f"{int(entry.get('impressions', 0)):,}"
        ])

    if current_adset_name:
        rows.append(["", current_adset_name + " 총합", "", f"₩{adset_total['daily_budget']:,}", f"{(adset_total['roas'] / adset_total['count']):.2f}", f"{(adset_total['ctr'] / adset_total['count']):.2f}%", f"₩{adset_total['spend']:,}", f"{adset_total['impressions']:,}"])
    if current_campaign_name:
        rows.append([current_campaign_name + " 총합", "", "", f"₩{campaign_total['daily_budget']:,}", f"{(campaign_total['roas'] / campaign_total['count']):.2f}", f"{(campaign_total['ctr'] / campaign_total['count']):.2f}%", f"₩{campaign_total['spend']:,}", f"{campaign_total['impressions']:,}"])

    worksheet.append_rows(rows, value_input_option="RAW")

    # 강조 색상 설정
    for idx, row in enumerate(rows, start=2):
        if "총합" in row[1]:
            format_cell_range(worksheet, f"B{idx}", CellFormat(backgroundColor=Color(0.5, 1.0, 0.5)))  # 초록색 강조
        if "총합" in row[0]:
            format_cell_range(worksheet, f"A{idx}", CellFormat(backgroundColor=Color(1.0, 1.0, 0.5)))  # 노란색 강조

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

class ProgressTracker:
    def __init__(self, checkpoint_file='progress_checkpoint.json'):
        self.checkpoint_file = checkpoint_file
        self.progress = self.load_progress()

    def load_progress(self):
        if os.path.exists(self.checkpoint_file):
            try:
                with open(self.checkpoint_file, 'r') as f:
                    return json.load(f)
            except:
                return self.get_initial_progress()
        return self.get_initial_progress()

    def get_initial_progress(self):
        return {
            'campaign_index': 0,
            'adset_index': 0,
            'ad_index': 0,
            'processed_data': []
        }

    def save_progress(self, campaign_idx, adset_idx, ad_idx, processed_data):
        self.progress.update({
            'campaign_index': campaign_idx,
            'adset_index': adset_idx,
            'ad_index': ad_idx,
            'processed_data': processed_data
        })
        with open(self.checkpoint_file, 'w') as f:
            json.dump(self.progress, f)

    def clear_progress(self):
        if os.path.exists(self.checkpoint_file):
            os.remove(self.checkpoint_file)
        self.progress = self.get_initial_progress()

# 메인 실행
if __name__ == "__main__":
    try:
        initialize_api()
        print("[INFO] Fetching active campaigns...")

        api_call_manager = APICallManager(max_calls_per_second=1, cooldown_time=300)
        all_data = []

        campaigns = fetch_with_retry(fetch_active_campaigns, ad_account_id=AD_ACCOUNT_ID, call_manager=api_call_manager)

        for campaign in campaigns:
            try:
                campaign_name = campaign[Campaign.Field.name]
                print(f"[INFO] Processing campaign: {campaign_name}")

                adsets = fetch_with_retry(fetch_active_adsets, campaign_id=campaign[Campaign.Field.id], call_manager=api_call_manager)
                
                for adset in adsets:
                    try:
                        adset_name = adset[AdSet.Field.name]
                        daily_budget = adset.get(AdSet.Field.daily_budget, 0)

                        ads = fetch_with_retry(fetch_active_ads, adset_id=adset[AdSet.Field.id], call_manager=api_call_manager)
                        
                        for ad in ads:
                            try:
                                ad_name = ad[Ad.Field.name]
                                print(f"[INFO] Processing ad: {ad_name}")

                                insights = fetch_with_retry(fetch_ad_insights, ad_id=ad[Ad.Field.id], call_manager=api_call_manager)
                                
                                for insight in insights:
                                    roas = insight.get("purchase_roas", [{}])[0].get("value", 0)
                                    ctr = insight.get("ctr", 0)
                                    spend = insight.get("spend", 0)
                                    impressions = insight.get("impressions", 0)

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
                            except Exception as e:
                                print(f"[WARNING] Error processing ad {ad_name}: {e}")
                                continue
                    except Exception as e:
                        print(f"[WARNING] Error processing adset {adset_name}: {e}")
                        continue
            except Exception as e:
                print(f"[WARNING] Error processing campaign {campaign_name}: {e}")
                continue

        if all_data:  # 수집된 데이터가 있는 경우에만 저장 시도
            print(f"[INFO] Saving {len(all_data)} records to Google Sheets...")
            try:
                fetch_with_retry(save_to_google_sheets, data=all_data)
                print("[INFO] Data saved successfully!")
            except Exception as e:
                print(f"[ERROR] Failed to save data to Google Sheets: {e}")
                # 백업 파일로 데이터 저장
                backup_file = f'facebook_ads_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
                with open(backup_file, 'w', encoding='utf-8') as f:
                    json.dump(all_data, f, ensure_ascii=False, indent=2)
                print(f"[INFO] Data backed up to {backup_file}")
        else:
            print("[WARNING] No data collected to save")

    except Exception as e:
        print(f"[ERROR] An error occurred: {e}")
        if all_data:  # 에러 발생 시에도 수집된 데이터가 있다면 백업
            backup_file = f'facebook_ads_data_error_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
            with open(backup_file, 'w', encoding='utf-8') as f:
                json.dump(all_data, f, ensure_ascii=False, indent=2)
            print(f"[INFO] Data backed up to {backup_file}")
        sys.exit(1)