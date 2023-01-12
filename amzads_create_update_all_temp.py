import requests
import requests_cache
import time
import csv
import sys
from requests.exceptions import HTTPError
from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import date

requests_cache.install_cache('amz_api_cache', backend='sqlite', expire_after=180)

#------------------------------------------------------------------------------
def get_access_token(REFRESH_TOKEN):
    data = {
        'grant_type': 'refresh_token',
        'client_id': client_id,
        'refresh_token': refresh_token,
        'client_secret': client_secret,
    }
    r = requests.post('https://api.amazon.com/auth/o2/token', data=data)
    r_json = r.json()
    return r_json["access_token"]

def csv_to_excel (csv_file, xl_file):
    csv_data = []
    with open(dl) as file_obj:
        reader = csv.reader(file_obj)
        for row in reader:
            csv_data.append(row)
        
        workbook = Workbook()
        sheet = workbook.active
        for row in csv_data:
            sheet.append(row)
        workbook.save(seo_report)

#------------------------------------------------------------------------------

dl = sys.argv[1]
#dl = r"C:\Users\anerl\Desktop\Amz ads files\ph2\export (37) - test.xlsx"

try:
    if 'csv' in dl:
        seo_report = f'{dl}'.replace('csv', 'xlsx')
        csv_to_excel(dl, seo_report)
    elif 'xlsx' in dl:
        seo_report = dl
except:
    print('File type must be csv or xlsx')

#------------------------------------------------------------------------------

wb = load_workbook(filename = seo_report)
ws = wb.worksheets[0]
max_row = ws.max_row
if ws[f'A{max_row}'].value == None:
    max_row = max_row - 1 

start_date = date.today().strftime('%Y%m%d')

campaigns = {}

for row_num, row in enumerate(ws.iter_rows(min_row=2, max_row=max_row ), start=2):
    sku = str(ws[f'A{row_num}'].value).strip()
    attribute_set = ws[f'C{row_num}'].value.strip()
    size = str(ws[f'E{row_num}'].value).strip()
    common_name =  ws[f'W{row_num}'].value.strip() 
    portfolio_id = ws[f'X{row_num}'].value   
    daily_budget = ws[f'Y{row_num}'].value
    ad_default_bid = ws[f'Z{row_num}'].value
    close_match_bid = ws[f'AA{row_num}'].value
    loose_match_bid = ws[f'AB{row_num}'].value
    substitutes_bid = ws[f'AC{row_num}'].value
    complements_bid = ws[f'AD{row_num}'].value

    ad_group_name = sku + ' - ' + size
    campaign_name = 'NS | ' + attribute_set + ' - ' + common_name

    
    if campaign_name in campaigns:
        campaigns[f'{campaign_name}']['sku_dict'][f'{sku}'] = [ad_group_name, ad_default_bid, {'close_match':[close_match_bid], 'loose_match':[loose_match_bid], 'complements':[complements_bid], 'substitutes':[substitutes_bid]}]
        
    else:
        campaigns[f'{campaign_name}'] = {
            'sku_dict': {sku: [ad_group_name, ad_default_bid, {'close_match':[close_match_bid], 'loose_match':[loose_match_bid], 'complements':[complements_bid], 'substitutes':[substitutes_bid]}]}, 
            'daily_budget': daily_budget,
            'attribute_set': attribute_set,
            'portfolio_id': portfolio_id,
            'campaign_name': campaign_name
            }
        
#------------------------------------------------------------------------------
#REQUEST CREDENTIALS

redirect_uri = 'https://fabhabitat.eu/packageui/advertisementAmazon.php'
client_id = 'amzn1.application-oa2-client.a3bff1cf1aba4d94bdef2badeae76dff'
client_secret = '445e7f6087acb91524ac14419205be302951e70693b8f5116995d0adda53a32e'
refresh_token =  'Atzr|IwEBIDu92De331amDsD5Uc5tTVPXSh35iiRjYyETM6Kmj-ozabz35WMaJrHOuDW-MhAGICKJ5YVTt9Wm3gqjntgJxMr3nRPnVN01-X43u6vkEtrxrHC7mycx_iiCGa4RD2hQm3uE8y_kj6YpR3BtLFsYEdguN0x5K6IxNxshVIJu2MRw_Rzad1wzRYT7UUhKJpkrXH8QQb8KIhDyvOg_grbWcPGPJJP-8hFCeAYDnjdUA9OA9XNcJXn7lC3xBAXZUfX-7VPIDkPyWadSBBpEvAQYejhZoOh6klrjsViOBVO4BhBC0xGdK8Etmb0Zb5bFpZvtO4e8XUhC-8KWJH8g4gxL0T286q6XsBZuhKHwFw8r_1J4NzCmbQAJMi0lwlipGMDq3RL0m2KL0cXWrIkUZ9IgJkTN-qJgayeAixloT7EMe4TyzGP8NMOzhUTGUpESwW6Q_mme7296rfsy_7d7ej3CjLKlWJG5bKK4UTaZ6LbcKNZVAQ'
profile_id = '2302325434421224'
access_token = get_access_token(refresh_token)

headers = {
    'Amazon-Advertising-API-ClientId': client_id,
    'Amazon-Advertising-API-Scope': profile_id,
    'Authorization': f'Bearer {access_token}',
}
 
#------------------------------------------------------------------------------
campaign_count = 0
retries = 3
campaign_url = 'https://advertising-api.amazon.com/v2/sp/campaigns'
adGroup_url = 'https://advertising-api.amazon.com/v2/sp/adGroups'
productAd_url = 'https://advertising-api.amazon.com/v2/sp/productAds'
targets_url = 'https://advertising-api.amazon.com/v2/sp/targets'

#CAMPAIGNS
for campaign_name, campaign_dict in campaigns.items():
    daily_budget = campaign_dict['daily_budget']
    portfolio_id = campaign_dict['portfolio_id']
    data =[{
       'portfolioId': portfolio_id,
       'name': campaign_name,
       'tags': {},
       'campaignType': 'sponsoredProducts',
       'targetingType': 'auto',
       'state': 'paused',
       'dailyBudget': daily_budget,
       'startDate': start_date,
       'endDate': None,
       'bidding': {'strategy': 'legacyForSales'} # = Dynamic down bids only
    }]
    s = requests.Session()
    for num in range(retries):
        try:
            r = s.post(campaign_url, headers=headers, json=data)
            r.raise_for_status()
            campaign_count+=1
            status_code = r.status_code
            r_json = r.json()
            campaign_id = r_json[0]['campaignId']
            campaign_dict['campaign_id'] = campaign_id
            print(f'Campaign {campaign_count} creation request status: {status_code}')
            time.sleep(1)
            break
        except HTTPError as exc:
            code = exc.r.status_code
            if code in [429, 500, 502, 503, 504]:
                time.sleep(3)
                continue
            raise
        except:
            status_code = r.status_code
            campaign_count+=1
            s.close()
            sys.exit(f'Error code: {status_code} while trying to create campaign {campaign_count}.\n{campaign_count-1} campaign(s) created.\n0 ad groups created.\n0 product ads created.\n0 product targeting expressions updated.')
    
    product_ad_count = 0    
    ad_group_count = 0    
    ad_group_id_list = []
    #--------------------------------------------------------
    #AD GROUPS
    for sku, attributes in campaign_dict['sku_dict'].items():
        ad_group_name = attributes[0]
        ad_default_bid = attributes[1]
        r_sku = str(sku)
        data =[{
           'name': ad_group_name,
           'campaignId': campaign_id,
           'defaultBid': ad_default_bid,
           'state': 'enabled'
        }]
        
        for num in range(retries):
            try:
                r =  s.post(adGroup_url, headers=headers, json=data)
                r.raise_for_status()
                ad_group_count+=1
                status_code = r.status_code
                r_json = r.json()
                ad_group_id = r_json[0]['adGroupId']
                #attributes.append(ad_group_id)         #idx 3
                #ad_group_id_list.append(str(ad_group_id))
                print(f'Ad group {ad_group_count} creation request status: {status_code}')
                time.sleep(1)
                break
            except HTTPError as exc:
                code = exc.r.status_code
                if code in [429, 500, 502, 503, 504]:
                    time.sleep(3)
                    continue
                raise
            except:
                status_code = r.status_code
                s.close()
                ad_group_count+=1
                sys.exit(f'Error code: {status_code} with Ad group {ad_group_count} creation request, in Campaign {campaign_count}.\n{campaign_count} campaign(s) created.\n{ad_group_count-1} ad groups created.\n0 product ads created.\n0 product targeting expressions updated.')
        
        
        #--------------------------------------------------------
        #PRODUCT ADS
        #ad_group_id = attributes[3]
        data =[{
           'campaignId': campaign_id,
           'adGroupId': ad_group_id,
           'sku': r_sku,
           'state': 'enabled'
        }]
        
        for num in range(retries):
            try:
                r =  s.post(productAd_url, headers=headers, json=data)
                r.raise_for_status()
                product_ad_count+=1
                status_code = r.status_code
                print(f'Product ad {product_ad_count} creation request status: {status_code}')
                time.sleep(1)
                break
            except HTTPError as exc:
                code = exc.r.status_code
                if code in [429, 500, 502, 503, 504]:
                    time.sleep(3)
                    continue
                raise
            except:
                status_code = r.status_code
                s.close()
                product_ad_count+=1
                sys.exit(f'Error code: {status_code} with Product ad {product_ad_count} creation request, in Campaign {campaign_count}, Ad group {ad_group_count}.\n{campaign_count} campaign(s) created.\n{ad_group_count} ad groups created.\n{product_ad_count-1} product ads created.\n0 product targeting expressions updated.')

        
        #------------------------------------------------------------------------------        
        #QUERY TARGETING
        ids_for_query = str(ad_group_id)
        query = {'stateFilter': 'enabled', }
        
        for num in range(retries):
            try:
                r = s.get(targets_url, params=query, headers=headers)
                r_json = r.json()
                r.raise_for_status()
                status_code = r.status_code
                print(f'Targets id request status: {status_code}')
                break
            except HTTPError as exc:
                code = exc.r.status_code
                if code in [429, 500, 502, 503, 504]:
                    print(f'Received error code {code} for target id lookup request. Trying again')
                    time.sleep(3)
                    continue
                raise
            except:
                status_code = r.status_code
                s.close()
                sys.exit(f'Error code: {status_code} with Targeting ids query, in Campaign {campaign_count}, Ad group {ad_group_count}.\n{campaign_count} campaign(s) created.\n{ad_group_count} ad groups created.\n{product_ad_count} product ads created.\n0 product targeting expressions updated.')

        
        #------------------------------------------------------------------------------            
        #UPDATE TARGETING        
        target_count = 0
        for target in r_json:
            if ad_group_id == target['adGroupId']:
                exp_type = target['expression'][0]['type']
                target_id = target['targetId']
                if exp_type == 'queryHighRelMatches': 
                    bid = attributes[2]['close_match'][0]
                elif exp_type == 'queryBroadRelMatches': 
                    bid = attributes[2]['loose_match'][0]
                elif exp_type == 'asinAccessoryRelated': 
                    bid = attributes[2]['complements'][0]
                elif exp_type == 'asinSubstituteRelated':
                    bid = attributes[2]['substitutes'][0]
                
                data =[{
                   'targetId': target_id,
                   'bid': bid
                }]
                
                for num in range(retries):
                    try:
                        r = s.put(targets_url, headers=headers, json=data)
                        r.raise_for_status()
                        status_code = r.status_code
                        target_count+=1
                        print(f'Ad group {ad_group_count} targeting expression {target_count} update status: {status_code}')
                        time.sleep(1)
                        break
                    except HTTPError as exc:
                        code = exc.r.status_code
                        if code in [429, 500, 502, 503, 504]:
                            time.sleep(3)
                            print(f'Received error code {code} for targeting update request. Trying again')
                            continue
                        raise
                    except:
                        status_code = r.status_code
                        target_count+=1
                        s.close()
                        sys.exit(f'Error code: {status_code} with updating target {target_count}, in Campaign {campaign_count}, Ad group {ad_group_count}.\n{campaign_count} campaign(s) created.\n{ad_group_count} ad groups created.\n{product_ad_count} product ads created.')



s.close()
















