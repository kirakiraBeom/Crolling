import requests
import time
import hmac
import hashlib
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import urllib.parse
import gspread_formatting as gsf
import math

# 네이버 API 설정
NAVER_CLIENT_KEY = '0100000000258e1a9e2d626b6e10c80471c41a9834842b9bf3303ce02b6f2de748ae513a1c'
NAVER_CLIENT_SECRET = 'AQAAAAAljhqeLWJrbhDIBHHEGpg0+/lKFc+RcMcyGJUgVScNeg=='
NAVER_CUSTOMER_ID = '2902933'
BASE_URL = 'https://api.searchad.naver.com'

# Google Sheets API 설정
GOOGLE_SHEET_ID = '15q1Gqr3iyQM8DM6rjmPWyYO92r-1jo5ESV4DrLltbS4'
JSON_KEY_FILE = r'C:\selenium\json\google_sheets_key.json'

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

# 활성화된 캠페인 검색
def fetch_campaigns():
    uri = "/ncc/campaigns"
    url = f"{BASE_URL}{uri}"
    headers = get_header("GET", uri)

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    campaigns = response.json()

    return [campaign for campaign in campaigns if campaign.get('status') == 'ELIGIBLE']

# 광고그룹 검색
def fetch_adgroups(campaign_id):
    uri = "/ncc/adgroups"
    url = f"{BASE_URL}{uri}"
    headers = get_header("GET", uri)
    params = {"nccCampaignId": campaign_id}

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    adgroups = response.json()

    return [adgroup for adgroup in adgroups if adgroup.get('status') == 'ELIGIBLE']

# 광고소재 검색
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
            reference_data = creative.get("referenceData", {})
            product_name = reference_data.get("productName", "Unknown Product")

            result.append({
                "nccAdId": creative.get("nccAdId"),
                "adgroupName": creative.get("nccAdgroupId"),
                "productName": product_name,
            })
    return result

# 광고소재 성과 검색
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
    yesterday = (datetime.now() - timedelta(days=1)).strftime('%y.%m.%d')
    sheet.clear()
    gsf.format_cell_range(sheet, f'A1:Z{sheet.row_count}', gsf.cellFormat(backgroundColor=gsf.color(1, 1, 1)))
    sheet.append_row([
        f"캠페인 이름({yesterday})", "광고그룹 이름", "광고소재 이름", "노출수", "클릭수", "클릭률(CTR)(%)", "평균클릭비용(CPC)(원)",
        "총비용(원)", "전환수", "전환율(%)", "전환매출액(원)", "광고수익률(%)", "전환당비용(원)"
    ])
    gsf.format_cell_range(sheet, 'A1:M1', gsf.cellFormat(textFormat=gsf.textFormat(bold=True, fontSize=12)))

    # 데이터를 한번에 추가
    rows = []
    prev_campaign_name = ""
    prev_adgroup_name = ""
    for entry in data.get("data", []):
        campaign_name = entry.get("campaignName", "")
        adgroup_name = entry.get("adgroupName", "")

        # 중복 처리: 캠페인 이름, 광고그룹 이름
        if campaign_name == prev_campaign_name:
            campaign_name = ""
        else:
            prev_campaign_name = entry.get("campaignName", "")

        if adgroup_name == prev_adgroup_name:
            adgroup_name = ""
        else:
            prev_adgroup_name = entry.get("adgroupName", "")

        rows.append([
            campaign_name,  # 캠페인 이름
            adgroup_name,  # 광고그룹 이름
            entry.get("productName", ""),  # 광고소재 이름
            f"{entry.get('impCnt', 0):,}",  # 노출수
            f"{entry.get('clkCnt', 0):,}",  # 클릭수
            round((entry.get('clkCnt', 0) / entry.get('impCnt', 1) * 100), 2) if entry.get('impCnt', 0) > 0 else 0,  # 클릭률(CTR)
            f"{int(round((entry.get('salesAmt', 0) / entry.get('clkCnt', 1)), 0)):,}" if entry.get("clkCnt", 0) > 0 else 0,  # 평균클릭비용(CPC)
            f"{entry.get('salesAmt', 0):,}",  # 총비용
            f"{entry.get('ccnt', 0):,}",  # 전환수
            round((entry.get("ccnt", 0) / entry.get("clkCnt", 1) * 100), 2) if entry.get("clkCnt", 0) > 0 else 0,  # 전환율(%)
            f"{entry.get('convAmt', 0):,}",  # 전환매출액
            round((entry.get("convAmt", 0) / entry.get("salesAmt", 1) * 100), 2) if entry.get("salesAmt", 0) > 0 else 0,  # 광고수익률(%)
            f"{int(round((entry.get('salesAmt', 0) / entry.get('ccnt', 1)), 0)):,}" if entry.get("ccnt", 0) > 0 else 0  # 전환당비용(원)
        ])
    sheet.append_rows(rows, value_input_option="RAW")

    # 셀 서식 설정 (가로 및 세로 정렬을 가운데로 설정)
    center_alignment_format = gsf.cellFormat(
        horizontalAlignment="CENTER",
        verticalAlignment="MIDDLE",
        textFormat=gsf.textFormat(bold=False)
    )
    gsf.format_cell_range(sheet, f'A1:M{len(rows) + 1}', center_alignment_format)
    
    # 1열 글씨체 굵게 설정
    bold_column_a_format = gsf.cellFormat(
        textFormat=gsf.textFormat(bold=True),
        horizontalAlignment="CENTER",
        verticalAlignment="MIDDLE"
    )
    gsf.format_cell_range(sheet, f'A1:A{len(rows) + 1}', bold_column_a_format)
    
    purple_format = gsf.cellFormat(
        backgroundColor=gsf.color(0.6, 0.4, 0.8),
        textFormat=gsf.textFormat(bold=True)
    )
    yellow_format = gsf.cellFormat(
        backgroundColor=gsf.color(1, 1, 0),
        textFormat=gsf.textFormat(bold=True)
    )
    green_format = gsf.cellFormat(
        backgroundColor=gsf.color(0, 1, 0),
        textFormat=gsf.textFormat(bold=True)
    )
    default_text_format = gsf.textFormat(bold=False)

    # 서식 설정을 한번에 처리
    yellow_cells = []
    green_cells = []
    purple_cells = []
    default_cells = []

    for idx, row in enumerate(rows, start=2):
        if "총합" in row[0]:
            for col_idx, cell_value in enumerate(row, start=1):
                if cell_value != "":
                    yellow_cells.append(f"{chr(64 + col_idx)}{idx}")
        elif "총합" in row[1]:
            for col_idx, cell_value in enumerate(row, start=1):
                if cell_value != "":
                    green_cells.append(f"{chr(64 + col_idx)}{idx}")
        else:
            default_cells.append(f"A{idx}:M{idx}")

    # 전체 캠페인 총합 행 추가 및 서식 지정
    if len(rows) > 0:
        total_stat = {
            "campaignName": "전체 캠페인 총합",
            "adgroupName": "",
            "productName": "",
            "impCnt": sum(int(row[3].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]),
            "clkCnt": sum(int(row[4].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]),
            "ctr": round((sum(int(row[4].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) / sum(int(row[3].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) * 100), 2) if sum(int(row[3].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) > 0 else 0,
            "cpc": f"{int(round((sum(int(row[7].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) / sum(int(row[4].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1])), 0)):,}" if sum(int(row[4].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) > 0 else 0,
            "salesAmt": sum(int(row[7].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]),
            "ccnt": sum(int(row[8].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]),
            "crto": round((sum(int(row[8].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) / sum(int(row[4].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) * 100), 2) if sum(int(row[4].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) > 0 else 0,
            "convAmt": sum(int(row[10].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]),
            "ror": round((sum(int(row[10].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) / sum(int(row[7].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) * 100), 2) if sum(int(row[7].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) > 0 else 0,
            "cpConv": f"{int(round((sum(int(row[7].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) / sum(int(row[8].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1])), 0)):,}" if sum(int(row[8].replace(",", "")) for row in rows if "총합" not in row[0] and "총합" not in row[1]) > 0 else 0
        }
        total_row = [
            total_stat["campaignName"], total_stat["adgroupName"], total_stat["productName"],
            f"{total_stat['impCnt']:,}", f"{total_stat['clkCnt']:,}", total_stat['ctr'], total_stat['cpc'],
            f"{total_stat['salesAmt']:,}", f"{total_stat['ccnt']:,}", total_stat['crto'], f"{total_stat['convAmt']:,}",
            total_stat['ror'], total_stat['cpConv']
        ]
        sheet.append_row(total_row, value_input_option="RAW")
        for col_idx, cell_value in enumerate(total_row, start=1):
            if cell_value != "":
                purple_cells.append(f"{chr(64 + col_idx)}{len(rows) + 2}")

    if purple_cells:
        gsf.format_cell_ranges(sheet, [(cell, purple_format) for cell in purple_cells])
    if yellow_cells:
        gsf.format_cell_ranges(sheet, [(cell, yellow_format) for cell in yellow_cells])
    if green_cells:
        gsf.format_cell_ranges(sheet, [(cell, green_format) for cell in green_cells])
    if default_cells:
        gsf.format_cell_ranges(sheet, [(cell, gsf.cellFormat(textFormat=default_text_format)) for cell in default_cells])

if __name__ == "__main__":
    try:
        print("[INFO] Fetching active campaigns...")
        # 활성화된 캠페인 검색
        campaigns = fetch_campaigns()

        # 모든 광고소재 통계 데이터를 저장할 리스트
        all_creative_stats = []

        # 캠페인을 순환하며 데이터를 검색
        for campaign in campaigns:
            print(f"[INFO] Processing campaign: {campaign['name']}")

            # 광고그룹 검색
            adgroups = fetch_adgroups(campaign['nccCampaignId'])
            for adgroup in adgroups:
                print(f"[INFO] Processing adgroup: {adgroup['name']}")

                # 광고소재 검색
                creatives = fetch_creatives_by_adgroup(adgroup['nccAdgroupId'])
                creative_ids = [creative["nccAdId"] for creative in creatives]

                if creative_ids:
                    # 광고소재 성과 검색
                    creative_stats = fetch_creative_stats(creative_ids)

                    # 성과 데이터를 정리하여 저장
                    for stat in creative_stats.get("data", []):
                        stat["campaignName"] = campaign["name"]  # 캠페인 이름 추가
                        stat["adgroupName"] = adgroup["name"]  # 광고그룹 이름 추가
                        stat["productName"] = next(
                            (creative["productName"] for creative in creatives if creative["nccAdId"] == stat["id"]),
                            "Unknown Product"
                        )
                        all_creative_stats.append(stat)

        # 광고그룹별로 데이터를 그룹화 및 정리
        grouped_data = {}
        for stat in all_creative_stats:
            campaign_name = stat["campaignName"]
            adgroup_name = stat["adgroupName"]
            key = (campaign_name, adgroup_name)

            if key not in grouped_data:
                grouped_data[key] = []
            grouped_data[key].append(stat)

        final_stats = []
        prev_campaign_name = ""
        prev_adgroup_name = ""

        for (campaign_name, adgroup_name), stats in grouped_data.items():
            for stat in stats:
                # 캠페인 이름 중복 제거
                if campaign_name == prev_campaign_name:
                    stat["campaignName"] = ""
                else:
                    prev_campaign_name = campaign_name

                # 광고그룹 이름 중복 제거
                if adgroup_name == prev_adgroup_name:
                    stat["adgroupName"] = ""
                else:
                    prev_adgroup_name = adgroup_name

                final_stats.append(stat)

            # 광고그룹별 총합 추가
            if len(stats) > 1:
                total_stat = {
                    "campaignName": "",  # 총합에서는 캠페인 이름을 비워둠
                    "adgroupName": f"{adgroup_name} 총합",
                    "productName": "",
                    "impCnt": sum(stat.get("impCnt", 0) for stat in stats),
                    "clkCnt": sum(stat.get("clkCnt", 0) for stat in stats),
                    "ctr": round((sum(stat.get("clkCnt", 0) for stat in stats) / sum(stat.get("impCnt", 1) for stat in stats) * 100), 2) if sum(stat.get("impCnt", 0) for stat in stats) > 0 else 0,
                    "cpc": f"{int(round((sum(stat.get("salesAmt", 0) for stat in stats) / sum(stat.get("clkCnt", 1) for stat in stats)), 0)):,}" if sum(stat.get("clkCnt", 0) for stat in stats) > 0 else 0,
                    "salesAmt": sum(stat.get("salesAmt", 0) for stat in stats),
                    "ccnt": sum(stat.get("ccnt", 0) for stat in stats),
                    "crto": round((sum(stat.get("ccnt", 0) for stat in stats) / sum(stat.get("clkCnt", 1) for stat in stats) * 100), 2) if sum(stat.get("clkCnt", 0) for stat in stats) > 0 else 0,
                    "convAmt": sum(stat.get("convAmt", 0) for stat in stats),
                    "ror": round((sum(stat.get("convAmt", 0) for stat in stats) / sum(stat.get("salesAmt", 1) for stat in stats) * 100), 2) if sum(stat.get("salesAmt", 0) for stat in stats) > 0 else 0,
                    "cpConv": f"{int(round((sum(stat.get("salesAmt", 0) for stat in stats) / sum(stat.get("ccnt", 1) for stat in stats)), 0)):,}" if sum(stat.get("ccnt", 0) for stat in stats) > 0 else 0
                }
                final_stats.append(total_stat)

        # 캠페인별 총합 추가
        updated_stats = []
        current_campaign = ""
        campaign_group_stats = []

        for stat in final_stats:
            if stat["campaignName"] and current_campaign and current_campaign != stat["campaignName"]:
                if len(campaign_group_stats) > 1:
                    campaign_total_stat = {
                        "campaignName": f"{current_campaign} 총합",
                        "adgroupName": "",
                        "productName": "",
                        "impCnt": sum(s.get("impCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                        "clkCnt": sum(s.get("clkCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                        "ctr": round((sum(s.get("clkCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("impCnt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) * 100), 2) if sum(s.get("impCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                        "cpc": f"{int(round((sum(s.get("salesAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("clkCnt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", ""))), 0)):,}" if sum(s.get("clkCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                        "salesAmt": sum(s.get("salesAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                        "ccnt": sum(s.get("ccnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                        "crto": round((sum(s.get("ccnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("clkCnt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) * 100), 2) if sum(s.get("clkCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                        "convAmt": sum(s.get("convAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                        "ror": round((sum(s.get("convAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("salesAmt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) * 100), 2) if sum(s.get("salesAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                        "cpConv": f"{int(round((sum(s.get("salesAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("ccnt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", ""))), 0)):,}" if sum(s.get("ccnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                    }
                    updated_stats.append(campaign_total_stat)
                campaign_group_stats = []

            current_campaign = stat["campaignName"] or current_campaign
            campaign_group_stats.append(stat)
            updated_stats.append(stat)

        if len(campaign_group_stats) > 1:
            campaign_total_stat = {
                "campaignName": f"{current_campaign} 총합",
                "adgroupName": "",
                "productName": "",
                "impCnt": sum(s.get("impCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                "clkCnt": sum(s.get("clkCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                "ctr": round((sum(s.get("clkCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("impCnt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) * 100), 2) if sum(s.get("impCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                "cpc": f"{int(round((sum(s.get("salesAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("clkCnt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", ""))), 0)):,}" if sum(s.get("clkCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                "salesAmt": sum(s.get("salesAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                "ccnt": sum(s.get("ccnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                "crto": round((sum(s.get("ccnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("clkCnt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) * 100), 2) if sum(s.get("clkCnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                "convAmt": sum(s.get("convAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")),
                "ror": round((sum(s.get("convAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("salesAmt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) * 100), 2) if sum(s.get("salesAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
                "cpConv": f"{int(round((sum(s.get("salesAmt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) / sum(s.get("ccnt", 1) for s in campaign_group_stats if "총합" not in s.get("adgroupName", ""))), 0)):,}" if sum(s.get("ccnt", 0) for s in campaign_group_stats if "총합" not in s.get("adgroupName", "")) > 0 else 0,
            }
            updated_stats.append(campaign_total_stat)

        # Google Sheets에 저장
        print("[INFO] Saving all stats to Google Sheets...")
        save_to_google_sheets({"data": updated_stats})

        print("[INFO] Completed successfully!")

    except Exception as e:
        print(f"[ERROR] {e}")
