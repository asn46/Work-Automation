import win32com.client as win32
from datetime import date
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as gcl


#----------------------------------------------------------------------
#FILE LOCATIONS

seo_report = r'C:\Users\anerl\Desktop\Amz ads files\m2 product seo report - test nantucket 3x5.xlsx'

save_path = r'C:\Users\anerl\Desktop\Amz ads files\test_auto_creation.xlsx'

#----------------------------------------------------------------------
#PULL DATA FROM SEO REPORT

wb = load_workbook(filename = seo_report)
seo_ws =  wb.worksheets[0]

sku =  seo_ws['A2'].value
collection = seo_ws['C2'].value
product_name = seo_ws['W2'].value  
size = seo_ws['E2'].value
start_date = date.today().strftime('%Y%m%d')
portfolio_id = seo_ws['X2'].value         #needs to be added to SEO report - adjust cell references once added
daily_budget = seo_ws['Y2'].value         #needs to be added to SEO report - adjust cell references once added
ad_default_bid = seo_ws['Z2'].value       #needs to be added to SEO report - adjust cell locations once added
 

campaign_name = f'NS | {collection}' + ' - ' + f'{product_name}' 
ad_group_name = f'{sku}' + ' - ' + f'{size}'
upload_sheet_name = 'Sheet1'

#----------------------------------------------------------------------
#CAMPAIGN/AD GROUP/PRODUCT AD SHEET CREATION

upload_sheet = wb.create_sheet(upload_sheet_name)


#creating list of headers & a few values common to all rows
default = {
    'Product': 'Sponsored Products',
    'Entity': None,
    'Operation': 'Create', 
    'Campaign Id': f'{campaign_name}', 
    'Ad Group Id': None, 
    'Portfolio Id': None,
    'Ad Id': None, 
    'Keyword Id': None, 
    'Product Targeting Id': None,
    'Campaign Name': None, 
    'Ad Group Name': None, 
    'Start Date': None,
    'End Date': None, 
    'Targeting Type': None, 
    'State': 'enabled',
    'Daily Budget': None, 
    'SKU': None, 
    'ASIN': None,
    'Ad Group Default Bid': None, 
    'Bid': None, 
    'Keyword Text': None,
    'Match Type': None, 
    'Bidding Strategy': None, 
    'Placement': None, 
    'Percentage': None, 
    'Product Targeting Expression': None
    }

col_num = 1

for header, value in default.items():
    col_letter = gcl(col_num)
    upload_sheet[f'{col_letter}1'] = header
    if value != None:
        upload_sheet[f'{col_letter}2'] = value
        upload_sheet[f'{col_letter}3'] = value
        upload_sheet[f'{col_letter}4'] = value
    col_num+=1    


#values specific to creating campaign - row 2
upload_sheet['B2'] = 'Campaign'
upload_sheet['F2'] = f'{portfolio_id}'
upload_sheet['J2'] = f'{campaign_name}'
upload_sheet['L2'] = f'{start_date}'
upload_sheet['N2'] = 'AUTO'
upload_sheet['P2'] = f'{daily_budget}'
upload_sheet['W2'] = 'Fixed Bid'

#values specific to creating ad group - row 3
upload_sheet['B3'] = 'Ad Group'
upload_sheet['E3'] = f'{ad_group_name}'  
upload_sheet['K3'] = f'{ad_group_name}'
upload_sheet['S3'] = f'{ad_default_bid}'

#values specific to creating product ad - row 4
upload_sheet['B4'] = 'Product Ad' 
upload_sheet['E4'] = f'{ad_group_name}' 
upload_sheet['Q4'] = sku  


wb.remove(seo_ws)
wb.save(save_path)


#opens file to calculate formulas
excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(save_path)
workbook.Save()
workbook.Close()
excel.Quit()



   

