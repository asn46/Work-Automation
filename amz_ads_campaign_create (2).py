#import win32com.client as win32
from datetime import date
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as gcl
import sys
import csv
import os

#----------------------------------------------------------------------
#IN CASE DOWNLOAD IS CSV

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

#----------------------------------------------------------------------
#FILE LOCATIONS FROM THE ARGUMENTS BEING PASSED

dl = sys.argv[1]

try:
    if 'csv' in dl:
        seo_report = f'{dl}'.replace('csv', 'xlsx')
        csv_to_excel(dl, seo_report)
    elif 'xlsx' in dl:
        seo_report = dl
except:
    print('File type must be csv or xlsx')

#----------------------------------------------------------------------
#OPEN SEO REPORT & STORE VALUES


wb = load_workbook(filename = seo_report)
seo_ws =  wb.worksheets[0]
seo_max_row = seo_ws.max_row

if seo_ws[f'A{seo_max_row}'].value == None:
    seo_max_row = seo_max_row - 1


start_date = date.today().strftime('%Y%m%d')

campaigns = {}

for row_num, row in enumerate(seo_ws.iter_rows(min_row=2, max_row=seo_max_row ), start=2):
    
    #design = seo_ws[f'W{row_num}'].value
    sku = seo_ws[f'A{row_num}'].value
    size = seo_ws[f'E{row_num}'].value
    daily_budget = seo_ws[f'Y{row_num}'].value         
    ad_default_bid = seo_ws[f'Z{row_num}'].value 
    product_name = seo_ws[f'D{row_num}'].value
    common_name =  seo_ws[f'W{row_num}'].value
    attribute_set = seo_ws[f'C{row_num}'].value
    #portfolio_id = seo_ws['X2'].value   #actual
    portfolio_id = '129484909483720' #test - zen collection
    
    if common_name in campaigns:
        campaigns[f'{common_name}']['skus'][f'{sku}'] = [size, product_name, common_name, ad_default_bid]
        
    else:
        campaigns[f'{common_name}'] = {
            'skus': {sku: [size, product_name, common_name, ad_default_bid]}, 
            'daily budget': daily_budget,
            'attribute set': attribute_set,
            'portfolio id': portfolio_id
            }


headers = [
    'Product',
    'Entity',
    'Operation',
    'Campaign Id', 
    'Ad Group Id', 
    'Portfolio Id',
    'Ad Id', 
    'Keyword Id', 
    'Product Targeting Id',
    'Campaign Name', 
    'Ad Group Name', 
    'Start Date',
    'End Date', 
    'Targeting Type', 
    'State',
    'Daily Budget', 
    'SKU', 
    'ASIN',
    'Ad Group Default Bid', 
    'Bid', 
    'Keyword Text',
    'Match Type', 
    'Bidding Strategy', 
    'Placement', 
    'Percentage', 
    'Product Targeting Expression'
    ]

    
#------------------------------------------------------------------------
#CREATING UPLOAD SHEETs

current_directory = os.getcwd()
save_directory = os.path.join(current_directory, r'Campaign creation sheets')

if not os.path.exists(save_directory):
    os.makedirs(save_directory)          

for common_name, design_dict in campaigns.items(): 
    save_name = common_name.strip() + '.xlsx'
    save_path = os.path.join(save_directory, save_name)
    
    wb2 = Workbook()
    ws = wb2.worksheets[0]
    
    row_count = 2
    
    #headers
    for col_num, header in enumerate(headers, start=1):
        col_letter = gcl(col_num)
        ws[f'{col_letter}1'] = header
       
    #values specific to creating campaign
    #campaign_name = '(TEST3) NS | Zen Collection Rugs - Khanti - Teal' #test
    campaign_name = 'NS | ' + attribute_set + ' - ' + common_name
    
    ws[f'B{row_count}'] = 'Campaign'
    ws[f'F{row_count}'] = design_dict['portfolio id']
    ws[f'D{row_count}'] = campaign_name
    
    ws[f'J{row_count}'] = campaign_name
    ws[f'L{row_count}'] = start_date
    ws[f'N{row_count}'] = 'AUTO'
    ws[f'W{row_count}'] = 'Fixed Bid'
    ws[f'P{row_count}'] = design_dict['daily budget']
    
    
    for sku, attributes in design_dict['skus'].items():
        row_count+=1
        p_size = attributes[0]
        p_adb = attributes[3]
       
        ad_group_name = str(sku) + ' - ' + p_size
        ws[f'D{row_count}'] = campaign_name
        ws[f'B{row_count}'] = 'Ad Group'
        ws[f'E{row_count}'] = ad_group_name
        ws[f'K{row_count}'] = ad_group_name
        ws[f'S{row_count}'] = p_adb  
        
#values for creating product ad
        row_count+=1    #currently set to 1 product ad per sku
        ws[f'D{row_count}'] = campaign_name
        ws[f'B{row_count}'] = 'Product Ad'
        ws[f'E{row_count}'] = ad_group_name
        ws[f'Q{row_count}'] = sku
       
#values common to all rows
    for row_num in range(2, row_count+1):
        ws[f'A{row_num}'] = 'Sponsored Products'
        ws[f'C{row_num}'] = 'Create'
        ws[f'O{row_num}'] = 'paused'    
    
    wb2.save(save_path)

                     

#------------------------------------------------------------------------

#excel = win32.gencache.EnsureDispatch('Excel.Application')
#workbook = excel.Workbooks.Open(save_path)
#workbook.Save()
#workbook.Close()
#excel.Quit()




