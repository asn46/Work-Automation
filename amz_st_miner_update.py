from copy import deepcopy
from datetime import date, datetime, timedelta
from pprint import pprint, pformat
import gzip
import io
import json
import logging
import re
import requests
import sys
import time

# -------------------------------------------------------------------------
# NOTES:
# - checks name for NS and V1 prefix. If manual campaign prefix is V2, V3, etc. then need to update script
# -------------------------------------------------------------------------

n = datetime.now()
now = n.strftime('%Y.%m.%d_%H.%M.%S')
logging.basicConfig(filename=f'{now}_log.txt', filemode='a',
                    format='%(asctime)s,%(msecs)d %(name)s %(levelname)s: %(message)s', datefmt='%H:%M:%S',
                    level=logging.DEBUG)
logger = logging.getLogger('myLogger')


class Main():

    def __init__(self):
        self.redirect_uri = 'https://fabhabitat.eu/packageui/advertisementAmazon.php'
        self.client_id = 'amzn1.application-oa2-client.a3bff1cf1aba4d94bdef2badeae76dff'
        self.client_secret = '445e7f6087acb91524ac14419205be302951e70693b8f5116995d0adda53a32e'
        self.refresh_token = 'Atzr|IwEBIDu92De331amDsD5Uc5tTVPXSh35iiRjYyETM6Kmj-ozabz35WMaJrHOuDW-MhAGICKJ5YVTt9Wm3gqjntgJxMr3nRPnVN01-X43u6vkEtrxrHC7mycx_iiCGa4RD2hQm3uE8y_kj6YpR3BtLFsYEdguN0x5K6IxNxshVIJu2MRw_Rzad1wzRYT7UUhKJpkrXH8QQb8KIhDyvOg_grbWcPGPJJP-8hFCeAYDnjdUA9OA9XNcJXn7lC3xBAXZUfX-7VPIDkPyWadSBBpEvAQYejhZoOh6klrjsViOBVO4BhBC0xGdK8Etmb0Zb5bFpZvtO4e8XUhC-8KWJH8g4gxL0T286q6XsBZuhKHwFw8r_1J4NzCmbQAJMi0lwlipGMDq3RL0m2KL0cXWrIkUZ9IgJkTN-qJgayeAixloT7EMe4TyzGP8NMOzhUTGUpESwW6Q_mme7296rfsy_7d7ej3CjLKlWJG5bKK4UTaZ6LbcKNZVAQ'
        self.profile_id = '2302325434421224'
        self.entities_kw = {}
        self.entities_asin = {}
        self.auto_campaign_name = None
        self.manual_campaign_name = []
        self.has_asins = False
        self.has_keywords = False
        self.has_manual_asin_campaign = True
        self.has_manual_kw_campaign = True
        self.session = requests.Session()
        self.access_token = self.get_access_token()
        self.base_headers = {
            'Amazon-Advertising-API-ClientId': self.client_id,
            'Amazon-Advertising-API-Scope': self.profile_id,
            'Authorization': f'Bearer {self.access_token}'
        }

    def get_access_token(self):
        data = {
            'grant_type': 'refresh_token',
            'client_id': self.client_id,
            'refresh_token': self.refresh_token,
            'client_secret': self.client_secret,
        }
        r = self.session.post('https://api.amazon.com/auth/o2/token', data=data)
        r_json = r.json()
        return r_json["access_token"]


    def set_name(self):
        self.auto_campaign_name = input('Enter name of AUTO campaign to check for new search terms:\n')
        asin_campaign_name_to_confirm = self.auto_campaign_name + ' - ASIN'
        kw_campaign_name_to_confirm = self.auto_campaign_name + ' - KW'
        num2 = input(
            f'\n{asin_campaign_name_to_confirm}\n{kw_campaign_name_to_confirm}\n\nEnter 1 to confirm that the above are the correct names for the MANUAL campaigns to update. Or enter 2 if these are incorrect:\n')
        if num2 == 1 or num2 == '1':
            self.manual_campaign_name = [kw_campaign_name_to_confirm, asin_campaign_name_to_confirm]
        else:
            self.manual_campaign_name[0] = input('Enter name of MANUAL KW campaign to update:\n')
            self.manual_campaign_name[1] = input('Enter name of MANUAL ASIN campaign to update:\n')


    def store_campaign(self, campaign_dict):
        campaignId = campaign_dict['campaignId']
        # portfolioId = campaign_dict['portfolioId']
        portfolioId = '129484909483720'  # test # incomplete portfolio
        campaign = {}
        self.entities_kw[f'{campaignId}'] = [campaign, {}]
        log = pformat(self.entities_kw)
        logger.debug(
            f'Campaign stored. self.entities_kw now contains:\n{log}\n--------------------------------------------\n\n')


    def query_auto_campaign(self):
        name_to_search = self.auto_campaign_name
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spCampaign.v3+json'
        headers_['Content-Type'] = 'application/vnd.spCampaign.v3+json'
        C_query_url = r'https://advertising-api.amazon.com/sp/campaigns/list'
        data = {
            'stateFilter': {
                'include': [
                    'ENABLED', 'PAUSED'
                ]
            },
            'nameFilter': {
                'queryTermMatchType': 'EXACT_MATCH',
                'include': [
                    name_to_search
                ]
            }
        }
        r = self.session.post(url=C_query_url, headers=self.base_headers, json=data)
        log = pformat(data)
        logger.debug(f'Campaign query data posted:\n{log}\n--------------------------------------------')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'Campaign query response:\n{log2}\n--------------------------------------------\n\n')
        if len(r_json['campaigns']) > 1:
            sys.exit('More than 1 matching campaign retrieved from query. Rerun script with exact campaign name')
        if len(r_json['campaigns']) < 1:
            sys.exit('No campaigns with that name were found')
        for campaign_dict in r_json['campaigns']:
            self.store_campaign(campaign_dict)


    def store_ad_group(self, adGroups):
        for ad_group_dict in adGroups:
            adGroupId = ad_group_dict['adGroupId']
            campaignId = ad_group_dict['campaignId']
            name = ad_group_dict['name']
            ad_group = {'name': name, 'adGroupId': adGroupId}
            for campaign_id, entity_group in self.entities_kw.items():
                if campaign_id == f'{campaignId}':
                    entity_group[1][f'{adGroupId}'] = {'adGroupData': ad_group, 'keywords': {},
                                                       'targetingExpressions': {}, 'productAds': {}}
                    log = pformat(self.entities_kw)
                    logger.debug(
                        f'Ad group stored., self.entities_kw now contains:\n{log}\n--------------------------------------------\n\n')


    def query_ad_groups(self):
        ad_group_url = r'https://advertising-api.amazon.com/sp/adGroups/list'
        campaignIds = [str(campaignId) for campaignId in self.entities_kw]
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spAdGroup.v3+json'
        headers_['Content-Type'] = 'application/vnd.spAdGroup.v3+json'
        data = {
            "campaignIdFilter": {
                "include": campaignIds
            },
            "stateFilter": {
                "include": [
                    "ENABLED", "PAUSED"
                ]
            }
        }
        r = self.session.post(url=ad_group_url, headers=headers_, json=data)
        r_json = r.json()
        log = pformat(r_json)
        logger.debug(f'Ad group query response:\n{log}\n--------------------------------------------')
        adGroups = r_json['adGroups']
        self.store_ad_group(adGroups)


    def create_report_and_return_id(self):
        full_url = r'https://advertising-api.amazon.com/reporting/reports'
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
                "timeUnit": "SUMMARY"
            },
            "name": "spSearchTermReport",
            "startDate": f"{today - thirty_days}"
        }
        r = self.session.post(url=full_url, headers=self.base_headers, json=data)
        r_json = r.json()
        reportId = r_json['reportId']
        logger.debug(f'ReportID: {reportId}\n--------------------------------------------')
        return reportId


    def report_url(self):
        report_id = self.create_report_and_return_id()
        full_url = fr'https://advertising-api.amazon.com/reporting/reports/{report_id}'
        logger.debug(f'Report url: {full_url}\n--------------------------------------------\n')
        r = self.session.get(url=full_url, headers=self.base_headers)
        r_json = r.json()
        report_status = r_json['status']
        while report_status == 'PENDING' or report_status == 'PROCESSING':
            time.sleep(10)
            r = self.session.get(url=full_url, headers=self.base_headers)
            r_json = r.json()
            report_status = r_json['status']
            print(f'Report status: {report_status}\n')
        report_url = r_json['url']
        return report_url


    def dl_report(self):
        r_url = self.report_url()
        r = self.session.get(url=r_url, stream=True)
        gz_file = r.content

        with gzip.open(io.BytesIO(gz_file), 'rb') as f:
            file_content = f.read()
            decoded = file_content.decode('UTF-8')
            content_dict = json.loads(decoded)
            print('Report retrieved\n')
            return content_dict


    def is_asin(self, search_term_text):
        asin_prefix = "^b0"
        check = re.search(asin_prefix, search_term_text)
        if check:
            if len(search_term_text) == 10:
                return True
            else:
                return False
        else:
            return False


    def store_target_terms_from_report(self):
        report = self.dl_report()
        filtered_st_count = 0
        target_count = 0
        kw_count = 0
        for campaign_id, entity_group in self.entities_kw.items():
            # if kw_count==3: #TEST
            # break   #TEST
            for st_dict in report:
                if str(campaign_id) == str(st_dict['campaignId']):
                    if not st_dict['acosClicks7d'] == "" and st_dict['acosClicks7d'] is not None:
                        if st_dict['acosClicks7d'] < 40:
                            filtered_st_count += 1
                            campaignId = st_dict['campaignId']
                            search_term = st_dict['searchTerm']
                            print(f'Search term: {search_term}')
                            cpc = st_dict['costPerClick']
                            # acosClicks7d = st_dict["acosClicks7d"]
                            # acosClicks14d = st_dict["acosClicks14d"]
                            # adGroupName =  st_dict["adGroupName"]
                            adGroupId = st_dict['adGroupId']
                            # clicks = st_dict['clicks']
                            # impressions = st_dict['impressions']
                            if self.is_asin(search_term):
                                asin_target = {
                                    'expression': [
                                        {
                                            'type': 'ASIN_SAME_AS',
                                            'value': search_term
                                        }
                                    ],
                                    'campaignId': None,
                                    'expressionType': 'MANUAL',
                                    'state': 'ENABLED',
                                    'bid': float(cpc),
                                    'adGroupId': None
                                }
                                entity_group[1][f'{adGroupId}']['targetingExpressions'][f'{search_term}'] = asin_target
                                target_count += 1

                            else:
                                keyword_variations = []
                                exact_kw_bid = (float(cpc)) * 1.1
                                exact_kw = {
                                    'campaignId': None,
                                    'matchType': 'EXACT',
                                    'state': 'ENABLED',
                                    'bid': exact_kw_bid,
                                    'adGroupId': None,
                                    'keywordText': search_term
                                }
                                phrase_kw_bid = (float(cpc)) * 1.05
                                phrase_kw = {
                                    'campaignId': None,
                                    'matchType': 'PHRASE',
                                    'state': 'ENABLED',
                                    'bid': phrase_kw_bid,
                                    'adGroupId': None,
                                    'keywordText': search_term
                                }
                                keyword_variations.append(exact_kw)
                                keyword_variations.append(phrase_kw)
                                entity_group[1][f'{adGroupId}']['keywords'][f'{search_term}'] = keyword_variations
                                kw_count += 1

        if target_count > 0:
            self.has_asins = True
        if kw_count > 0:
            self.has_keywords = True
        print(f'Total filtered search terms found: {filtered_st_count}\n')
        print(f'{target_count} targeting expressions')
        print(f'{kw_count}*2 keywords\n')
        logger.debug(
            f'Total filtered search terms found: {filtered_st_count}\n--------------------------------------------')


    def create_asin_entity_dict_if_asins(self):
        if self.has_asins == True:
            self.entities_asin = deepcopy(self.entities_kw)


    def update_kw_dict_with_manual_campaign_id(self):
        name_kw = self.manual_campaign_name[0]
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spCampaign.v3+json'
        headers_['Content-Type'] = 'application/vnd.spCampaign.v3+json'
        C_query_url = r'https://advertising-api.amazon.com/sp/campaigns/list'
        data = {
            'stateFilter': {
                'include': [
                    'ENABLED',
                    'PAUSED'
                ]
            },
            'nameFilter': {
                'queryTermMatchType': 'EXACT_MATCH',
                'include': [
                    name_kw
                ]
            }
        }
        r = self.session.post(url=C_query_url, headers=self.base_headers, json=data)
        log = pformat(data)
        logger.debug(f'Manual kw campaign query data posted:\n{log}\n--------------------------------------------')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'Manual kw campaign query response:\n{log2}\n--------------------------------------------\n\n')
        if len(r_json['campaigns']) > 1:
            logger.debug('More than 1 matching manual KW campaign found from name entered. Did not perfom update on manual KW campaign.\n--------------------------------------------\n\n')
            self.has_manual_kw_campaign = False
            return
        elif len(r_json['campaigns']) < 1:
            logger.debug('No Manual KW campaigns found for this auto campaign\n--------------------------------------------\n\n')
            self.has_manual_kw_campaign = False
            return
        for campaign_dict in r_json['campaigns']:
            manual_campaign_id_kw = campaign_dict['campaignId']
        if self.has_keywords:
            for campaign_id, entity_group in self.entities_kw.copy().items():
                self.entities_kw[f'{manual_campaign_id_kw}'] = self.entities_kw[f'{campaign_id}']
                del self.entities_kw[f'{campaign_id}']
        else:
            print('self.has_keywords=False line 451')


    def update_asin_dict_with_manual_campaign_id(self):
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spCampaign.v3+json'
        headers_['Content-Type'] = 'application/vnd.spCampaign.v3+json'
        C_query_url = r'https://advertising-api.amazon.com/sp/campaigns/list'
        name_asin = self.manual_campaign_name[1]
        data = {
            'stateFilter': {
                'include': [
                    'ENABLED',
                    'PAUSED'
                ]
            },
            'nameFilter': {
                'queryTermMatchType': 'EXACT_MATCH',
                'include': [
                    name_asin
                ]
            }
        }
        r = self.session.post(url=C_query_url, headers=self.base_headers, json=data)
        log = pformat(data)
        logger.debug(f'Manual asin campaign query data posted:\n{log}\n--------------------------------------------')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'Manual asin campaign query response:\n{log2}\n--------------------------------------------\n\n')
        if len(r_json['campaigns']) > 1:
            logger.debug('More than 1 matching manual asin campaign found from name entered. Did not perfom update on manual asin campaign.')
            self.has_manual_asin_campaign = False
            return
        if len(r_json['campaigns']) < 1:
            if 'NS' in name_asin:
                name_asin = name_asin.replace('NS', 'V1')
                data['nameFilter']['include'] = [name_asin]
                r = self.session.post(url=C_query_url, headers=headers_, json=data)
                log = pformat(data)
                logger.debug(f'Attempt # 2 manual asin campaign query data posted:\n{log}\n--------------------------------------------')
                r_json = r.json()
                log2 = pformat(r_json)
                logger.debug(f'Attempt # 2 manual asin campaign query response:\n{log2}\n--------------------------------------------\n\n')
                if len(r_json['campaigns']) < 1:
                    logger.debug('No Manual asin campaigns found for this auto campaign')
                    self.has_manual_asin_campaign = False
                    return
            else:
                logger.debug('No Manual asin campaigns found for this auto campaign')
                self.has_manual_asin_campaign = False
                return
        for campaign_dict in r_json['campaigns']:
            manual_campaign_id_asin = campaign_dict['campaignId']
        if self.has_asins:
            for campaign_id, entity_group in self.entities_asin.copy().items():
                self.entities_asin[f'{manual_campaign_id_asin}'] = self.entities_asin[f'{campaign_id}']
                del self.entities_asin[f'{campaign_id}']
        else:
            print('self.has_asins=False line 485')


    def update_kw_dict_with_manual_ad_group(self):
        if self.has_manual_kw_campaign == False:
            logger.debug(f'self.has_manual_kw_campaign == False ; update_kw_dict_with_manual_ad_group not executed\n--------------------------------------------\n\n')
            return
        for campaign_id, entity_group in self.entities_kw.items():
            c_id_kw = str(campaign_id)
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spAdGroup.v3+json'
        headers_['Content-Type'] = 'application/vnd.spAdGroup.v3+json'
        C_query_url = r'https://advertising-api.amazon.com/sp/adGroups/list'
        data = {
            'campaignIdFilter': {
                'include': [
                    c_id_kw
                ]
            },
            'stateFilter': {
                'include': [
                    'ENABLED',
                    'PAUSED'
                ]
            },
        }
        r = self.session.post(url=C_query_url, headers=self.base_headers, json=data)
        log = pformat(data)
        logger.debug(f'Manual kw ad group query data posted:\n{log}\n--------------------------------------------')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'Manual kw ad group query response:\n{log2}\n--------------------------------------------\n\n')
        if 'adGroups' in r_json:
            for ad_group in r_json['adGroups']:
                query_ad_group_name = ad_group['name']
                query_ad_group_id = ad_group['adGroupId']
                for campaign_id, entity_group in self.entities_kw.copy().items():
                    for ad_group_id, attributes in entity_group[1].copy().items():
                        stored_ad_group_name = attributes['adGroupData']['name']
                        if stored_ad_group_name == query_ad_group_name:
                            self.entities_kw[f'{campaign_id}'][1][f'{query_ad_group_id}'] = \
                            self.entities_kw[f'{campaign_id}'][1][f'{ad_group_id}']
                            del self.entities_kw[f'{campaign_id}'][1][f'{ad_group_id}']
        else:
            print('no ad groups found, line 502')


    def update_asin_dict_with_manual_ad_group(self):
        if self.has_manual_asin_campaign == False:
            logger.debug(f'self.has_manual_asin_campaign == False; update_asin_dict_with_manual_ad_group not executed\n--------------------------------------------\n\n')
            return
        for campaign_id, entity_group in self.entities_asin.items():
            c_id_asin = str(campaign_id)
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spAdGroup.v3+json'
        headers_['Content-Type'] = 'application/vnd.spAdGroup.v3+json'
        C_query_url = r'https://advertising-api.amazon.com/sp/adGroups/list'
        data = {
            'campaignIdFilter': {
                'include': [
                    c_id_asin
                ]
            },
            'stateFilter': {
                'include': [
                    'ENABLED',
                    'PAUSED'
                ]
            },
        }
        r = self.session.post(url=C_query_url, headers=self.base_headers, json=data)
        log = pformat(data)
        logger.debug(f'Manual asin ad group query data posted:\n{log}\n--------------------------------------------')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'Manual asin ad group query response:\n{log2}\n--------------------------------------------\n\n')
        if 'adGroups' in r_json:
            for ad_group in r_json['adGroups']:
                query_ad_group_name = ad_group['name']
                query_ad_group_id = ad_group['adGroupId']
                for campaign_id, entity_group in self.entities_asin.copy().items():
                    for ad_group_id, attributes in entity_group[1].copy().items():
                        stored_ad_group_name = attributes['adGroupData']['name']
                        if stored_ad_group_name == query_ad_group_name:
                            self.entities_asin[f'{campaign_id}'][1][f'{query_ad_group_id}'] = \
                            self.entities_asin[f'{campaign_id}'][1][f'{ad_group_id}']
                            del self.entities_asin[f'{campaign_id}'][1][f'{ad_group_id}']
        else:
            print('no ad groups found, line 502')


    def check_existing_kw_and_del_duplicates(self):
        keyword_url = r'https://advertising-api.amazon.com/sp/keywords/list'
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spKeyword.v3+json'
        headers_['Content-Type'] = 'application/vnd.spKeyword.v3+json'
        if self.has_manual_kw_campaign and self.has_keywords:
            for campaign_id, entity_group in self.entities_kw.items():
                campaignId = str(campaign_id)
            data = {
                'campaignIdFilter': {
                    'include': [
                        campaignId
                    ]
                },
                'stateFilter': {
                    'include': [
                        'ENABLED',
                        'PAUSED'
                    ]
                }
            }
            r = self.session.post(url=keyword_url, headers=headers_, json=data)
            log = pformat(data)
            logger.debug(
                f'check_existing_kw_and_del_duplicates data posted:\n{log}\n--------------------------------------------')
            r_json = r.json()
            log2 = pformat(r_json)
            logger.debug(f'check_existing_kw_and_del_duplicates kw query response:\n{log2}\n')
            total_keywords_new = set()
            for campaign_id, entity_group in self.entities_kw.copy().items():
                for ad_group_id, attributes in entity_group[1].items():
                    for keyword_text, keyword_variation_list in attributes['keywords'].copy().items():
                        total_keywords_new.add(keyword_text)
                        for keyword in r_json['keywords']:
                            keyword_text_to_check = keyword['keywordText']
                            if keyword_text == keyword_text_to_check:
                                del self.entities_kw[f'{campaign_id}'][1][f'{ad_group_id}']['keywords'][f'{keyword_text}']
                                total_keywords_new.remove(keyword_text)
            if len(total_keywords_new) == 0:
                self.has_keywords =  False
        else:
            logger.debug('check_existing_kw_and_del_duplicates not executed\n-----------------------------------------\n\n')


    def check_existing_target_and_del_duplicates(self):
        if self.has_manual_asin_campaign and self.has_asins:
            keyword_url = r'https://advertising-api.amazon.com/sp/targets/list'
            headers_ = self.base_headers
            headers_['Accept'] = 'application/vnd.spTargetingClause.v3+json'
            headers_['Content-Type'] = 'application/vnd.spTargetingClause.v3+json'
            for campaign_id, entity_group in self.entities_asin.items():
                campaignId = str(campaign_id)
            data = {
                'campaignIdFilter': {
                    'include': [
                        campaignId
                    ]
                },
                'stateFilter': {
                    'include': [
                        'ENABLED',
                        'PAUSED'
                    ]
                }
            }
            r = self.session.post(url=keyword_url, headers=headers_, json=data)
            log = pformat(data)
            logger.debug(
                f'check_existing_target_and_del_duplicates data posted:\n{log}\n--------------------------------------------')
            r_json = r.json()
            log2 = pformat(r_json)
            logger.debug(f'check_existing_target_and_del_duplicates kw query response:\n{log2}\n')
            total_targets_new = set()
            for campaign_id, entity_group in self.entities_asin.copy().items():
                for ad_group_id, attributes in entity_group[1].items():
                    for target_text, target_dict in attributes['targetingExpressions'].copy().items():
                        total_targets_new.add(target_text)
                        for target in r_json['targetingClauses']:
                            target_text_to_check = target['expression'][0]['value']
                            if target_text == target_text_to_check:
                                del self.entities_asin[f'{campaign_id}'][1][f'{ad_group_id}']['targetingExpressions'][f'{target_text}']
                                total_targets_new.remove(target_text)
            if len(total_targets_new) == 0:
                self.has_asins = False
        else:
            logger.debug('check_existing_target_and_del_duplicates not executed\n-----------------------------------------\n\n')


    def create_keywords(self):
        if self.has_manual_kw_campaign and self.has_keywords:
            keywords_to_create = []
            for campaign_id, entity_group in self.entities_kw.items():
                for ad_group_id, attributes in entity_group[1].items():
                    for keyword_name, keyword_variation_list in attributes['keywords'].items():
                        for keyword_dict in keyword_variation_list:
                            keyword_dict['campaignId'] = campaign_id
                            keyword_dict['adGroupId'] = ad_group_id
                            keywords_to_create.append(keyword_dict)
            data = {'keywords': keywords_to_create}
            keyword_url = r'https://advertising-api.amazon.com/sp/keywords'
            headers_ = self.base_headers
            headers_['Accept'] = 'application/vnd.spKeyword.v3+json'
            headers_['Content-Type'] = 'application/vnd.spKeyword.v3+json'
            log = pformat(data)
            logger.debug(f'KW creation data posted:\n{log}\n--------------------------------------------')
            r = self.session.post(url=keyword_url, headers=headers_, json=data)
            r_json = r.json()
            log2 = pformat(r_json)
            logger.debug(f'KW creation response:\n{log2}\n')
            if len(r_json['keywords']['error'])>=1:
                print('Errors encountered while trying to add the following keywords:\n')
                for idx, keyword in enumerate(keywords_to_create):
                    k_name = keyword['expression'][0]['value']
                    for error in r_json['keywords']['error']:
                        if idx == error['index']:
                            error_to_print = pformat(error["errors"])
                            print(f'Name:{k_name}:\nError:{error_to_print}\n')
            else:
                print('No errors while adding keywords\n')

            if len(r_json['keywords']['success'])>=1:
                print('Following new keywords added successfully:')
                for idx, keyword in enumerate(keywords_to_create):
                    k_name = keyword['expression'][0]['value']
                    for success in r_json['keywords']['success']:
                        if idx == success['index']:
                            print(k_name)
                print('\n')
            else:
                print('No keywords added successfully')
        else:
            logger.debug('create_keywords not executed\n-----------------------------------------\n\n')


    def create_targeting_expressions(self):
        if self.has_manual_asin_campaign and self.has_asins:
            targeting_expressions_to_create = []
            for campaign_id, entity_group in self.entities_asin.items():
                for ad_group_id, attributes in entity_group[1].items():
                    for target_name, target_dict in attributes['targetingExpressions'].items():
                        target_dict['campaignId'] = campaign_id
                        target_dict['adGroupId'] = ad_group_id
                        targeting_expressions_to_create.append(target_dict)
            data = {'targetingClauses': targeting_expressions_to_create}
            target_url = r'https://advertising-api.amazon.com/sp/targets'
            headers_ = self.base_headers
            headers_['Accept'] = 'application/vnd.spTargetingClause.v3+json'
            headers_['Content-Type'] = 'application/vnd.spTargetingClause.v3+json'
            log = pformat(data)
            logger.debug(f'Target creation data posted:\n{log}\n--------------------------------------------')
            r = self.session.post(url=target_url, headers=headers_, json=data)
            r_json = r.json()
            log2 = pformat(r_json)
            logger.debug(f'Target creation response:\n{log2}\n')
            if len(r_json['targetingClauses']['error'])>=1:
                print('Errors encountered while trying to add the following asin targets:')
                for idx, target in enumerate(targeting_expressions_to_create):
                    t_name = target['expression'][0]['value']
                    for error in r_json['targetingClauses']['error']:
                        if idx == error['index']:
                            error_to_print = pformat(error["errors"])
                            print(f'Name:{t_name}:\nError:{error_to_print}\n')
            else:
                print('No errors while adding targets\n')

            if len(r_json['targetingClauses']['success'])>=1:
                print('Following new asin targets added successfully:')
                for idx, target in enumerate(targeting_expressions_to_create):
                    t_name = target['expression'][0]['value']
                    for success in r_json['targetingClauses']['success']:
                        if idx == success['index']:
                            print(t_name)
                print('\n')
            else:
                print('No asin targets added successfully')

        else:
            logger.debug('create_targeting_expressions not executed\n-----------------------------------------\n\n')


    def close_session(self):
        self.session.close()
        print('Script finished running')

# ----------------------------------------------------------------------------


if __name__ == '__main__':

    script = Main()
    script.set_name()
    script.query_auto_campaign()
    time.sleep(1)
    script.query_ad_groups()
    time.sleep(1)
    script.store_target_terms_from_report()
    time.sleep(1)
    script.create_asin_entity_dict_if_asins()
    script.update_kw_dict_with_manual_campaign_id()
    script.update_asin_dict_with_manual_campaign_id()
    script.update_kw_dict_with_manual_ad_group()
    script.update_asin_dict_with_manual_ad_group()
    if script.has_keywords:
        script.check_existing_kw_and_del_duplicates()
        script.create_keywords()
    if script.has_asins:
        script.check_existing_target_and_del_duplicates()
        script.create_targeting_expressions()
    script.close_session()