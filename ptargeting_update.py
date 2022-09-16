#for downloading and updating product targeting expressions (close match, loose, etc)
import win32com.client as win32
from openpyxl import load_workbook


#-------------------------------------------------------------

#[placeholder for code that downloads ptargeting sheet]

#-------------------------------------------------------------
#FILE LOCATIONS
seo_report = r'C:\Users\anerl\Desktop\Amz ads files\m2 product seo report - test nantucket 3x5.xlsx'
ptargeting_download = r"C:\Users\anerl\Desktop\Amz ads files\test_ptargeting_adj.xlsx"
save_location = r"C:\Users\anerl\Desktop\Amz ads files\ptargeting_upload.xlsx"

#------------------------------------------------------------
#DATA FROM SEO REPORT
wb = load_workbook(seo_report)

ws = wb.active

collection = ws['C2'].value
product_name = ws['W2'].value  

#campaign_name = 'NS | FH Home Flatweave Freeport twister - Budapest Auto (Secondary)' #TEST
campaign_name = f'NS | {collection}' + ' - ' + f'{product_name}' 

cmatch_bid = ws['AA2'].value
lmatch_bid = ws['AB2'].value
subs_bid = ws['AC2'].value

wb.close()

#------------------------------------------------------------
#UPDATING PRODUCT TARGETING

wb2 = load_workbook(ptargeting_download)
ws = wb2.worksheets[1]


cmatch_target = 'close-match'
lmatch_target = 'loose-match'
subs_target = 'substitutes'

ws_max_row = ws.max_row

for row_num, row in enumerate(ws.iter_rows(min_row=2, max_row=ws_max_row), start=2):
    row_list = []
    for cell in row:
        row_list.append(cell.value)
    if cmatch_target in row_list and campaign_name in row_list:
        ws[f'AB{row_num}'] = cmatch_bid
    elif lmatch_target in row_list and campaign_name in row_list:
        ws[f'AB{row_num}'] = lmatch_bid
    elif subs_target in row_list and campaign_name in row_list:
        ws[f'AB{row_num}'] = subs_bid
    else:
        pass
    

wb2.save(save_location)


#opens file to calculate formulas
excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(save_location)
workbook.Save()
workbook.Close()
excel.Quit()
