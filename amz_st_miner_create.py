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

# ----------------------------------------------
# NOTES:
# -
# ----------------------------------------------


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
        self.name = None
        self.has_asins = False
        self.has_keywords = False
        self.create_KW_entities =  True
        self.create_ASIN_entities = True
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
        self.name = str(input('Copy & paste AUTO campaign name from Ads Campaign Manager:\n'))


    def store_campaign(self, campaign_dict):
        campaignId = campaign_dict['campaignId']
        #portfolioId = campaign_dict['portfolioId']
        portfolioId = '129484909483720'  # test # incomplete portfolio
        c_name = campaign_dict['name']
        dynamicBidding = campaign_dict['dynamicBidding']
        startDate = str(date.today().strftime('%Y-%m-%d'))
        #budget = campaign_dict['budget']
        campaign = {
            'portfolioId': portfolioId,
            'name': c_name,
            'targetingType': 'MANUAL',
            'state': 'PAUSED',
            'dynamicBidding': dynamicBidding,
            'startDate': startDate,
            'budget': {
                'budgetType': 'DAILY',
                'budget': 5.00
            },
        }
        self.entities_kw[f'{campaignId}'] = [campaign, {}]
        log = pformat(self.entities_kw)
        logger.debug(
            f'Campaign stored. self.entities_kw now contains:\n{log}\n--------------------------------------------\n\n')


    def store_ad_group(self, adGroups):
        for ad_group_dict in adGroups:
            adGroupId = ad_group_dict['adGroupId']
            name = ad_group_dict['name']
            defaultBid = ad_group_dict['defaultBid']
            campaignId = ad_group_dict['campaignId']
            ad_group = {
                'campaignId': None,
                'name': name,
                'state': 'ENABLED',
                'defaultBid': defaultBid
            }
            for campaign_id, entity_group in self.entities_kw.items():
                if campaign_id == f'{campaignId}':
                    entity_group[1][f'{adGroupId}'] = {'adGroupData': ad_group, 'keywords': {},
                                                       'targetingExpressions': {}, 'productAds': {}}
                    log = pformat(self.entities_kw)
                    logger.debug(
                        f'Ad group stored., self.entities_kw now contains:\n{log}\n--------------------------------------------\n\n')


    def store_product_ad(self, productAds):
        for product_ad_dict in productAds:
            adId = product_ad_dict['adId']
            sku = product_ad_dict['sku']
            adGroupId = product_ad_dict['adGroupId']
            campaignId = product_ad_dict['campaignId']
            product_ad = {
                'campaignId': None,
                'state': 'ENABLED',
                'sku': sku,
                'adGroupId': None
            }
            for campaign_id, entity_group in self.entities_kw.items():
                if campaign_id == f'{campaignId}':
                    for ad_group_id, attributes in entity_group[1].items():
                        if ad_group_id == f'{adGroupId}':
                            attributes['productAds'][f'{adId}'] = product_ad
                            log = pformat(self.entities_kw)
                            logger.debug(
                                f'Product ad stored, self.entities_kw now contains:\n{log}\n--------------------------------------------\n\n')

    def query_campaigns(self):
        name_to_search = self.name
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
        log = pformat(data)
        logger.debug(f'Campaign query data posted:\n{log}\n--------------------------------------------')
        try:
            r = self.session.post(url=C_query_url, headers=self.base_headers, json=data)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            logger.debug(f'Encountered the following error while attempting campaign query:\n{e}\n------------------------------------------------')
            raise SystemExit(f'Encountered the following error while attempting campaign query:\n{e}')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'Campaign query response:\n{log2}\n--------------------------------------------\n\n')
        if len(r_json['campaigns']) > 1:
            sys.exit('More than 1 matching campaign retrieved from query. Rerun script with exact campaign name')
        elif len(r_json['campaigns']) < 1:
            sys.exit('No campaigns with that name were found')
        for campaign_dict in r_json['campaigns']:
            self.store_campaign(campaign_dict)


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
        log = pformat(data)
        logger.debug(f'Ad group query data posted:\n{log}\n--------------------------------------------')
        try:
            r = self.session.post(url=ad_group_url, headers=headers_, json=data)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            logger.debug(f'Encountered the following error while attempting ad group query:\n{e}\n------------------------------------------------')
            raise SystemExit(f'Encountered the following error while attempting ad group query:\n{e}')
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
                            if self.is_asin(search_term) == True:
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
                                k1_bid = (float(cpc)) * 1.1
                                keyword1 = {
                                    'campaignId': None,
                                    'matchType': 'EXACT',
                                    'state': 'ENABLED',
                                    'bid': k1_bid,
                                    'adGroupId': None,
                                    'keywordText': search_term
                                }
                                k2_bid = (float(cpc)) * 1.05
                                keyword2 = {
                                    'campaignId': None,
                                    'matchType': 'PHRASE',
                                    'state': 'ENABLED',
                                    'bid': k2_bid,
                                    'adGroupId': None,
                                    'keywordText': search_term
                                }
                                keyword_variations.append(keyword1)
                                keyword_variations.append(keyword2)
                                entity_group[1][f'{adGroupId}']['keywords'][f'{search_term}'] = keyword_variations
                                kw_count += 1

        if target_count > 0:
            self.has_asins = True
        if kw_count > 0:
            self.has_keywords = True
        print(f'Total filtered search terms found: {filtered_st_count}\n')
        print(f'{target_count} targeting expressions')
        print(f'{kw_count}*2 keywords')
        logger.debug(
            f'Total filtered search terms found: {filtered_st_count}\n--------------------------------------------')


    def update_entities_if_asins(self):
        if self.has_asins == True:
            self.entities_asin = deepcopy(self.entities_kw)


    def query_product_ads(self):
        campaignIds = [campaignId for campaignId in self.entities_kw]
        adGroupIds = []
        for campaign_id, entity_group in self.entities_kw.items():
            for ad_group_id in entity_group[1]:
                adGroupIds.append(ad_group_id)
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spProductAd.v3+json'
        headers_['Content-Type'] = 'application/vnd.spProductAd.v3+json'
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
        try:
            r = self.session.post(url=r'https://advertising-api.amazon.com/sp/productAds/list', headers=headers_, json=data)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            logger.debug(f'Encountered the following error while attempting product ad query:\n{e}\n------------------------------------------------')
            raise SystemExit(f'Encountered the following error while attempting product ad query:\n{e}')
        r_json = r.json()
        log = pformat(r_json)
        logger.debug(f'Product ad query response:\n{log}\n--------------------------------------------')
        productAds = r_json['productAds']
        self.store_product_ad(productAds)


    def create_manual_campaigns(self, entities, type):
        num = input(f'Enter 1 to confirm {type} campaign creation or enter 2 to not create:\n')
        if num == 2 or num == '2':
            if type == 'KW':
                self.create_KW_entities = False
            if type == 'ASIN':
                self.create_ASIN_entities = False
        elif num == 1 or num == '1':
            if type == 'KW':
                self.create_KW_entities = True
            if type == 'ASIN':
                self.create_ASIN_entities = True
        else:
            SystemExit('Invalid entry, need to enter 1 or 2. Rerun script')
        if type == 'KW':
            if self.create_KW_entities == False:
                logger.debug(f'create_manual_campaigns for {type} not executed\n')
                print(f'create_manual_campaigns for {type} not executed\n')
                return
        elif type == 'ASIN':
            if self.create_ASIN_entities == False:
                logger.debug(f'create_manual_campaigns for {type} not executed\n')
                print(f'create_manual_campaigns for {type} not executed\n')
                return
        campaign_url = r'https://advertising-api.amazon.com/sp/campaigns'
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spCampaign.v3+json'
        headers_['Content-Type'] = 'application/vnd.spCampaign.v3+json'
        campaigns_to_create = []
        for campaign_id, entity_group in entities.items():
            campaign = entity_group[0]
            campaigns_to_create.append(campaign)
        for idx, campaign in enumerate(campaigns_to_create):
            name = campaigns_to_create[idx]['name']
            if '|' in name:
                repl_idx = name.index('|')
                text_to_repl = name[:repl_idx]
                new_name = name.replace(text_to_repl, 'V1 ')
            else:
                start_str = name[0]
                new_name = name.replace(start_str, f'V1 {start_str}')
            new_name = new_name + ' - ' + type
            campaigns_to_create[idx]['name'] = new_name
        data = {'campaigns': campaigns_to_create}
        log = pformat(data)
        logger.debug(f'{type} Campaign creation data posted:\n{log}\n--------------------------------------------')
        try:
            r = self.session.post(url=campaign_url, headers=headers_, json=data)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            logger.debug(f'Encountered the following error while attempting campaign creation:\n{e}\n------------------------------------------------')
            raise SystemExit(f'Encountered the following error while attempting campaign creation:\n{e}')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'{type} Campaign creation response:\n{log2}\n')
        for campaign_dict_new in r_json['campaigns']['success']:
            id_new = campaign_dict_new['campaignId']
            idx_new = campaign_dict_new['index']
            for idx_old, (_id, entity_group) in enumerate(entities.copy().items()):
                if idx_old == idx_new:
                    entities[f'{id_new}'] = entities.get(f'{str(_id)}')
                    entities[f'{id_new}'][0] = 'campaign created successfully'
                    del entities[f'{str(_id)}']


    def create_ad_groups(self, entities, type):
        if type == 'KW':
            if self.create_KW_entities == False:
                return
        elif type == 'ASIN':
            if self.create_ASIN_entities == False:
                return
        old_ad_group_ids = []
        ad_groups_to_create = []
        for campaign_id, entity_group in entities.items():
            for ad_group_id, attributes in entity_group[1].items():
                attributes['adGroupData']['campaignId'] = str(campaign_id)
                ad_groups_to_create.append(attributes['adGroupData'])
                old_ad_group_ids.append([campaign_id, ad_group_id])
        data = {'adGroups': ad_groups_to_create}
        ad_group_url = r'https://advertising-api.amazon.com/sp/adGroups'
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spAdGroup.v3+json'
        headers_['Content-Type'] = 'application/vnd.spAdGroup.v3+json'
        log = pformat(data)
        logger.debug(f'{type} ad group creation data posted:\n{log}\n--------------------------------------------')
        try:
            r = self.session.post(url=ad_group_url, headers=headers_, json=data)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            logger.debug(f'Encountered the following error while attempting ad group creation:\n{e}\n------------------------------------------------')
            print(f'Encountered the following error while attempting ad group creation:\n{e}\n')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'{type} ad group creation response:\n{log2}\n')
        for ad_group_dict_new in r_json['adGroups']['success']:
            ad_group_id_new = ad_group_dict_new['adGroupId']
            idx_new = ad_group_dict_new['index']
            for idx_old, id_pair in enumerate(old_ad_group_ids):
                if idx_new == idx_old:
                    campaign_id = id_pair[0]
                    ad_group_id_old = id_pair[1]
                    entities[f'{campaign_id}'][1][f'{ad_group_id_new}'] = entities[f'{campaign_id}'][1][f'{ad_group_id_old}']
                    del entities[f'{campaign_id}'][1][f'{ad_group_id_old}']


    def create_product_ads(self, entities, type):
        if type == 'KW':
            if self.create_KW_entities == False:
                return
        elif type == 'ASIN':
            if self.create_ASIN_entities == False:
                return
        old_product_ad_ids = []
        product_ads_to_create = []
        for campaign_id, entity_group in entities.items():
            for ad_group_id, attributes in entity_group[1].items():
                for product_ad_id, product_ad in attributes['productAds'].items():
                    product_ad['campaignId'] = str(campaign_id)
                    product_ad['adGroupId'] = str(ad_group_id)
                    product_ads_to_create.append(product_ad)
        data = {'productAds': product_ads_to_create}
        headers_ = self.base_headers
        headers_['Accept'] = 'application/vnd.spProductAd.v3+json'
        headers_['Content-Type'] = 'application/vnd.spProductAd.v3+json'
        log = pformat(data)
        logger.debug(f'{type} product ad creation data posted:\n{log}\n--------------------------------------------')
        try:
            r = self.session.post(url=r'https://advertising-api.amazon.com/sp/productAds', headers=headers_, json=data)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            logger.debug(f'Encountered the following error while attempting product ad creation:\n{e}\n------------------------------------------------')
            print(f'Encountered the following error while attempting product ad creation:\n{e}\n')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'{type} product ad creation response:\n{log2}\n')


    def create_keywords(self):
        if self.create_KW_entities == False:
            return
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
        try:
            r = self.session.post(url=keyword_url, headers=headers_, json=data)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            logger.debug(f'Encountered the following error while attempting KW creation:\n{e}\n------------------------------------------------')
            print(f'Encountered the following error while attempting KW creation:\n{e}\n')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'KW creation response:\n{log2}\n')
        if len(r_json['keywords']['error']) >= 1:
            print('Errors encountered while trying to add the following keywords:\n')
            for idx, keyword in enumerate(keywords_to_create):
                k_name = keyword['expression'][0]['value']
                for error in r_json['keywords']['error']:
                    if idx == error['index']:
                        error_to_print = pformat(error["errors"])
                        print(f'Name:{k_name}:\nError:{error_to_print}\n')


    def create_targeting_expressions(self):
        if self.create_ASIN_entities == False:
            return
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
        try:
            r = self.session.post(url=target_url, headers=headers_, json=data)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            logger.debug(f'Encountered the following error while attempting target creation:\n{e}\n------------------------------------------------')
            print(f'Encountered the following error while attempting target creation:\n{e}\n')
        r_json = r.json()
        log2 = pformat(r_json)
        logger.debug(f'Target creation response:\n{log2}\n')
        if len(r_json['targetingClauses']['error']) >= 1:
            print('Errors encountered while trying to add the following asin targets:\n')
            for idx, target in enumerate(targeting_expressions_to_create):
                t_name = target['expression'][0]['value']
                for error in r_json['targetingClauses']['error']:
                    if idx == error['index']:
                        error_to_print = pformat(error["errors"])
                        print(f'Name:{t_name}:\nError:{error_to_print}\n')


    def close_session(self):
        self.session.close()
        print('Script finished running')



# ----------------------------------------------------------------------------


if __name__ == '__main__':

    script = Main()
    script.set_name()
    script.query_campaigns()
    time.sleep(1)
    script.query_ad_groups()
    time.sleep(1)
    script.store_target_terms_from_report()
    time.sleep(1)
    script.query_product_ads()
    time.sleep(1)
    script.update_entities_if_asins()
    time.sleep(1)
    if script.has_keywords == True:
        script.create_manual_campaigns(script.entities_kw, 'KW')
        time.sleep(2)
        script.create_ad_groups(script.entities_kw, 'KW')
        time.sleep(2)
        script.create_product_ads(script.entities_kw, 'KW')
        time.sleep(2)
        script.create_keywords()
    if script.has_asins == True:
        script.create_manual_campaigns(script.entities_asin, 'ASIN')
        time.sleep(2)
        script.create_ad_groups(script.entities_asin, 'ASIN')
        time.sleep(2)
        script.create_product_ads(script.entities_asin, 'ASIN')
        time.sleep(2)
        script.create_targeting_expressions()
    script.close_session()




