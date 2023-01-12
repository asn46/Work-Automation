from datetime import date, timedelta
from pprint import pprint
import gzip
import io
import json
import pandas as pd
import requests
import sys
import time
import logging

# --------------
id_list = []


# Functions----------
def get_access_token():
    data = {
        'grant_type': 'refresh_token',
        'client_id': client_id,
        'refresh_token': refresh_token,
        'client_secret': client_secret,
    }
    r = requests.post('https://api.amazon.com/auth/o2/token', data=data)
    r_json = r.json()
    return r_json["access_token"]


def create_report_and_return_id(session, timeUnit='SUMMARY'):
    endpoint = '/reporting/reports'
    full_url = base_url + endpoint
    today = date.today()
    today_str = today.strftime('%Y-%m-%d')
    thirty_days = timedelta(days=30)
    columns = [
        "impressions",
        "clicks",
        "costPerClick",
        "acosClicks7d",
        "acosClicks14d",
        "campaignId",
        "adGroupId",
        "adGroupName",
        "targeting",
        "searchTerm",
        "keywordType",
        "matchType",
        "keywordId",
    ]
    data = {
        "endDate": today_str,
        "configuration": {
            "adProduct": "SPONSORED_PRODUCTS",
            "columns": columns,
            "reportTypeId": "spSearchTerm",
            "format": "GZIP_JSON",
            "groupBy": [
                "searchTerm"
            ],
            "timeUnit": timeUnit
        },
        "name": "spSearchTermReport",
        "startDate": f"{today - thirty_days}"
    }
    r = session.post(url=full_url, headers=base_headers, json=data)
    r_json = r.json()
    reportId = r_json['reportId']
    # print(f'ReportID: {reportId}\n')
    return reportId


def report_url(session):
    report_id = create_report_and_return_id(session)
    print(f'Report ID is: {report_id}\nCreating report...\n')
    endpoint = fr'/reporting/reports/{report_id}'
    full_url = base_url + endpoint
    r = session.get(url=full_url, headers=base_headers)
    r_json = r.json()
    report_status = r_json['status']
    while report_status == 'PENDING' or report_status == 'PROCESSING':
        time.sleep(10)
        r = session.get(url=full_url, headers=base_headers)
        r_json = r.json()
        report_status = r_json['status']
        print(f'Report status: {report_status}\n')
    report_url = r_json['url']
    return report_url


def dl_report(session):
    r_url = report_url(session)
    headers_ = base_headers
    print(f'URL is: {r_url}')
    time.sleep(2)
    r = session.get(url=r_url, stream=True)
    gz_file = r.content

    with gzip.open(io.BytesIO(gz_file), 'rb') as f:
        file_content = f.read()
        decoded = file_content.decode('UTF-8')
        content_dict = json.loads(decoded)
        return content_dict


def targeted_search_terms_to_xlsx(session):
    name = str(input('Copy and paste campaign name from Amz Seller Central:\n'))
    query_campaign_info(name, 'Campaign', session)
    report = dl_report(session)
    search_terms = []
    for id in id_list:
        for st_dict in report:
            if str(id) == str(st_dict['campaignId']):
                if not st_dict['acosClicks14d'] == "" and st_dict['acosClicks14d'] is not None:
                    if st_dict['acosClicks14d'] <= 25:
                        campaignId = st_dict['campaignId']
                        text = st_dict['searchTerm']
                        matchType = st_dict['matchType']
                        keywordType = st_dict['keywordType']
                        targeting = st_dict['targeting']
                        cpc = st_dict['costPerClick']
                        acosClicks7d = st_dict["acosClicks7d"]
                        acosClicks14d = st_dict["acosClicks14d"]
                        adGroupName = st_dict["adGroupName"]
                        adGroupId = st_dict['adGroupId']
                        clicks = st_dict['clicks']
                        impressions = st_dict['impressions']
                        dict_to_add = {'campaignId': campaignId, 'text': text, 'matchType': matchType,
                                       'targeting': targeting, 'keywordType': keywordType, 'cpc': cpc,
                                       'acosClicks7d': acosClicks7d, 'acosClicks14d': acosClicks14d,
                                       'adGroupName': adGroupName, 'adGroupId': adGroupId, 'clicks': clicks,
                                       'impressions': impressions}
                        search_terms.append(dict_to_add)

    df = pd.DataFrame.from_dict(search_terms)
    df.to_excel('extracted_search_terms.xlsx')


def query_campaign_info(name_, entity_, session):
    name = name_
    entity = entity_
    if entity == 'Campaign':
        endpoint = '/sp/campaigns/list'
        headers_ = base_headers
        headers_['Accept'] = 'application/vnd.spCampaign.v3+json'
        headers_['Content-Type'] = 'application/vnd.spCampaign.v3+json'
        C_query_url = base_url + endpoint
        data = {
            'nameFilter': {
                'queryTermMatchType': 'EXACT_MATCH',
                'include': [
                    name
                ]
            },
            'stateFilter': {
                'include': [
                    'ENABLED'
                    # 'PAUSED'
                ]
            }
        }
        r = session.post(url=C_query_url, headers=headers_, json=data)
        r_json = r.json()
        print('campaign query:')
        pprint(r_json)
        print('\n')
        campaignId = r_json['campaigns'][0]['campaignId']
        id_list.append(campaignId)
        # if operation == 'Create':
        # store_for_campaign_creation(campaign_dict)


    elif entity == 'Portfolio':
        endpoint = '/v2/portfolios'
        url = base_url + endpoint
        params = {'portfolioNameFilter': name, 'portfolioStateFilter': 'enabled'}
        r = session.get(url, params=params)
        r_json = r.json()
        portfolioId = r_json[0]['portfolioId']
        headers_ = base_headers
        headers_['Accept'] = 'application/vnd.spCampaign.v3+json'
        headers_['Content-Type'] = 'application/vnd.spCampaign.v3+json'
        endpoint = '/sp/campaigns/list'
        C_query_url = base_url + endpoint
        data = {
            'portfolioIdFilter': {
                'include': [
                    str(portfolioId)
                ]
            },
            'stateFilter': {
                'include': [
                    'ENABLED',

                ]
            }
        }
        r = session.post(url=C_query_url, headers=headers_, json=data)
        r_json = r.json()

        # for campaign_dict in r_json['campaigns']:
        # campaignId = campaign_dict['campaignId']
        # id_list.append(campaignId)
        # if operation == 'Create':
        # store_for_campaign_creation(campaign_dict)


# ----------


redirect_uri = 'https://fabhabitat.eu/packageui/advertisementAmazon.php'
client_id = 'amzn1.application-oa2-client.a3bff1cf1aba4d94bdef2badeae76dff'
client_secret = '445e7f6087acb91524ac14419205be302951e70693b8f5116995d0adda53a32e'
refresh_token = 'Atzr|IwEBIDu92De331amDsD5Uc5tTVPXSh35iiRjYyETM6Kmj-ozabz35WMaJrHOuDW-MhAGICKJ5YVTt9Wm3gqjntgJxMr3nRPnVN01-X43u6vkEtrxrHC7mycx_iiCGa4RD2hQm3uE8y_kj6YpR3BtLFsYEdguN0x5K6IxNxshVIJu2MRw_Rzad1wzRYT7UUhKJpkrXH8QQb8KIhDyvOg_grbWcPGPJJP-8hFCeAYDnjdUA9OA9XNcJXn7lC3xBAXZUfX-7VPIDkPyWadSBBpEvAQYejhZoOh6klrjsViOBVO4BhBC0xGdK8Etmb0Zb5bFpZvtO4e8XUhC-8KWJH8g4gxL0T286q6XsBZuhKHwFw8r_1J4NzCmbQAJMi0lwlipGMDq3RL0m2KL0cXWrIkUZ9IgJkTN-qJgayeAixloT7EMe4TyzGP8NMOzhUTGUpESwW6Q_mme7296rfsy_7d7ej3CjLKlWJG5bKK4UTaZ6LbcKNZVAQ'
profile_id = '2302325434421224'
access_token = get_access_token()
base_headers = {
    'Amazon-Advertising-API-ClientId': client_id,
    'Amazon-Advertising-API-Scope': profile_id,
    'Authorization': f'Bearer {access_token}'
}

base_url = r'https://advertising-api.amazon.com'

# --------------

if __name__ == "__main__":
    with requests.Session() as s:
        targeted_search_terms_to_xlsx(s)

    print('Excel file created in current directory')



