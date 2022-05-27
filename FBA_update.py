from time import sleep
import gspread
from re import match as rm
from oauth2client.service_account import ServiceAccountCredentials
from freshdesk.v2.api import API
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

#--------------------------------------------------------------------


def FBA_ticket_check():
    a = API('fabhabitat2021.freshdesk.com', '7mcIyxGVhEvDhzyLDrV')
    FBA_orders = []
    added = []
    for ticket in a.tickets.list_tickets():
        
        fullstring = str(ticket)
        
        if 'FBA Inbound Shipment' in fullstring and not 'Bill' in fullstring:
            FBA_number =  fullstring[fullstring.find("(")+1 : fullstring.rfind(")")]
            substring1 = 'FBA Inbound Shipment Checked-In'
            substring2 = 'FBA Inbound Shipment Receiving'
            substring3 = 'FBA Inbound Shipment Received In-Full'
            datetime = str(ticket.created_at)
            date_match = rm(r'\d\d\d\d-\d\d-\d\d', datetime)
            date = date_match.group(0)
            
            if substring1 in fullstring:
                update = ['Reached FBA', 1]
            elif substring2 in fullstring:
                update = ['Being Received', 2]
            elif substring3 in fullstring:
                update = ['Closed', 3]
            else:
                raise Exception('There was an error [1].')
            
            if FBA_number not in added:
                print('Creating new fba_order dict...')
                fba_order = dict.fromkeys(['Name', 'Order', 'Status', 'Status_priority', 'Created Date', 'Ship Date', 'Update Date', 'Ship Qty', 'Rcv Qty'])
                fba_order['Name'] = FBA_number
                fba_order['Status'] = update[0]
                fba_order['Status_priority'] = update[1]
                fba_order['Update Date'] = date
                FBA_orders.append(fba_order)
                added.append(FBA_number)
            
            else:
                for fba_order in FBA_orders:
                    if fba_order['Name'] == FBA_number:
                        if update[1] > fba_order['Status_priority']:
                            fba_order['Status'] = update[0]
                            print('Updating existing fba_order dict...')
                        else:
                            print('Existing status takes priority, no changes made...')                        
    
    return FBA_orders

#-------------------------------------------------------------------------
not_updated = []

def SC_FBA_orders_update(playwright, FBA_orders):
    url = 'https://sellercentral.amazon.com/gp/ssof/shipping-queue.html/ref=xx_fbashipq_dnav_xx#fbashipment'
    context = playwright.chromium.launch_persistent_context(
        user_data_dir="C:\\Users\\aseem\\AppData\\Local\\Google\\Chrome\\User Data",
        headless=False,
        args=["--profile-directory=Default"],
    )
    page = context.new_page()
    page.goto(url)
    sleep(1)
    row = page.locator('id=content-row')

    for fba_order in FBA_orders:
        FBA_number = fba_order['Name']
        try:
            target_row = row.locator(':scope', has_text=f'{FBA_number}', timeout=6)
        except PlaywrightTimeoutError:
            global not_updated
            not_updated.append(FBA_number)
            FBA_orders.remove(fba_order)
            continue
        cell = target_row.locator('#numeric-cell >> nth=1')
        rcv_qty = target_row.locator('id=received_quantity_text').text_content()
        #ship_qty = cell.text_content()
        fba_order['Rcv Qty'] = rcv_qty
        print(f'Stored new Received Qty: {rcv_qty} for FBA order {FBA_number}')
    
    FBA_orders_updated = FBA_orders
    context.close()

    return FBA_orders_updated

#-----------------------------------------------------------------------
not_in_gsheet = []

def transfer_to_gsheet(FBA_orders):
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(secret.creds, scope)
    client = gspread.authorize(creds)
    
    with sync_playwright() as p:
        shipment_updates = SC_FBA_orders_update(p, FBA_orders)
    
    count = 0
    if shipment_updates:
        worksheet = client.open('FBA_test_sheet').sheet1
        transfer = []
        for fba_order in shipment_updates:
            FBA_number = fba_order['Name']
            status = fba_order['Status']
            rcv_qty = fba_order['Rcv Qty']
            update_date = fba_order['Update Date']
            try:
                shipment_row = worksheet.find(f'{FBA_number}').row
            except:
                not_in_gsheet.append(FBA_number)
                shipment_updates.remove(fba_order)
                continue
            
            if fba_order['Status'] == 'Reached FBA' or fba_order['Status'] == 'Being Received':
                col_letter = 'F'
            else:                      
                col_letter = 'G'
             
            shipment_rv = [{
                'range': f'C{shipment_row}', 
                'values': [[f'{status}']]
            }, {
                'range': f'{col_letter}{shipment_row}',
                'values': [[f'{update_date}']]
            }, {
                'range': f'J{shipment_row}',
                'values': [[f'{rcv_qty}']]
            }]
            
            for rv_dict in shipment_rv:
                transfer.append(rv_dict)
            count+=1
            
        worksheet.batch_update(transfer)

    if count == 1:
        print(f'-- {count} update was made --')
    else:
        print(f'-- {count} updates were made --')
    
#-----------------------------------------------------------------------
def main():
    FBA_orders = FBA_ticket_check()
    if not FBA_orders:
        print('-- No FBA update tickets found --')
        
    else:
        transfer_to_gsheet(FBA_orders)
        if not_updated:
            print(f'{not_updated} not found in Amz US FBA Shipments, check if they are EU')
            print('\n')
        if not_in_gsheet:
            print(f'{not_in_gsheet} not found in google sheet.')
            
#-----------------------------------------------------------------------
if __name__ == "__main__":
    main()
