from datetime import datetime, timedelta
from sp_api.api import ReportsV2
from sp_api.api import Finances
from sp_api.base.reportTypes import ReportType
from sp_api.base import Marketplaces
import requests
import pandas as pd
import time
import csv
import pytz
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
import io
import os
from io import BytesIO


#SOUTH-AFRICA
refresh_token = #ADD YOUR REFRESH TOKEN HERE-__-
usa_now_utc = datetime.now(pytz.utc)
# Calculate the date 48 hours earlier in UTC
usa_earlier_date_utc = usa_now_utc - timedelta(days=2)
# Define the start and end time for the entire earlier date in UTC
usa_data_start_time = datetime(usa_earlier_date_utc.year, usa_earlier_date_utc.month, usa_earlier_date_utc.day, 0, 0, 0, tzinfo=pytz.utc).isoformat()
usa_data_end_time = datetime(usa_earlier_date_utc.year, usa_earlier_date_utc.month, usa_earlier_date_utc.day, 23, 59, 59, tzinfo=pytz.utc).isoformat()
#south-africa
marketplace_southafrica = Marketplaces.ZA #You czn change this to your marketplace

#SOUTH-AFRICA CREDENTIALS =
client_id = #ADD YOUR CLIENT ID HERE-__-
client_secret = #ADD YOUR CLIENT SECRET HERE-__-

def finance_api(refresh_token, date_start_time, date_end_time, marketplace):
    credentials = dict(
    refresh_token= refresh_token, 
    lwa_app_id=client_id,  # From Seller Central, named CLIENT IDENTIFIER on website.
    lwa_client_secret=client_secret,  # From Seller Central, named CLIENT SECRET on website.
    )
    
    def fetch_financial_events(finances, date_start_time, date_end_time):
        financial_event_data = []
        next_token = None
        
        while True:
            # Make API call to list financial events
            if next_token:
                res = finances.list_financial_events(
                    PostedAfter=date_start_time,
                    PostedBefore=date_end_time,
                    NextToken=next_token
                )
            else:A
                res = finances.list_financial_events(
                    PostedAfter=date_start_time,
                    PostedBefore=date_end_time
                )
            
            # Extract payload from the response
            payload = res.payload
            
            # Check if FinancialEvents exist in payload
            if 'FinancialEvents' in payload:
                financial_events = payload['FinancialEvents']
                
                # Extract shipment events details
                if 'ShipmentEventList' in financial_events:
                    for shipment_event in financial_events['ShipmentEventList']:
                        amazon_order_id = shipment_event.get('AmazonOrderId', '')
                        seller_order_id = shipment_event.get('SellerOrderId', '')
                        marketplace_name = shipment_event.get('MarketplaceName', '')
                        posted_date = shipment_event.get('PostedDate', '')
                        
                        for item in shipment_event.get('ShipmentItemList', []):
                            seller_sku = item.get('SellerSKU', '')
                            order_item_id = item.get('OrderItemId', '')
                            quantity_shipped = item.get('QuantityShipped', 0)
                            
                            charges = {}
                            currency_code = None
                            
                            # Process item charges if available
                            if 'ItemChargeList' in item:
                                for charge in item['ItemChargeList']:
                                    charge_type = charge.get('ChargeType', '')
                                    charge_amount = charge.get('ChargeAmount', {}).get('CurrencyAmount', 0)
                                    currency_code = charge.get('ChargeAmount', {}).get('CurrencyCode', '')
                                    
                                    # Add charge amount to the corresponding charge type column
                                    charges[charge_type] = charge_amount
                            
                            # Process item fees if available
                            if 'ItemFeeList' in item:
                                for fee in item['ItemFeeList']:
                                    fee_type = fee.get('FeeType', '')
                                    fee_amount = fee.get('FeeAmount', {}).get('CurrencyAmount', 0)
                                    if currency_code is None:
                                        currency_code = fee.get('FeeAmount', {}).get('CurrencyCode', '')
                                    
                                    # Add fee amount to the corresponding fee type column
                                    charges[fee_type] = fee_amount
                            
                            # Process promotions if available
                            if 'PromotionList' in item:
                                for promotion in item['PromotionList']:
                                    promotion_type = promotion.get('PromotionType', '')
                                    promotion_amount = promotion.get('PromotionAmount', {}).get('CurrencyAmount', 0)
                                    if currency_code is None:
                                        currency_code = promotion.get('PromotionAmount', {}).get('CurrencyCode', '')
                                    
                                    # Add promotion amount to the corresponding promotion type column
                                    charges[promotion_type] = promotion_amount
                            
                            # Append the extracted data to the list
                            financial_event_data.append({
                                'AmazonOrderId': amazon_order_id,
                                'SellerOrderId': seller_order_id,
                                'MarketplaceName': marketplace_name,
                                'PostedDate': posted_date,
                                'SellerSKU': seller_sku,
                                'OrderItemId': order_item_id,
                                'QuantityShipped': quantity_shipped,
                                'CurrencyCode': currency_code,
                                **charges
                            })
            
            # Check for next token to continue pagination
            next_token = payload.get('NextToken')
            if not next_token:
                break  # No more pages
            
        return financial_event_data
        
    if marketplace == refresh_token:
        finances = Finances(credentials=credentials)
    else:
        finances = Finances(credentials=credentials, marketplace=marketplace) 
    
    # Fetch financial events data with pagination
    financial_event_data = fetch_financial_events(finances, date_start_time, date_end_time)
    
    # Convert the list to a DataFrame
    financial_events_df = pd.DataFrame(financial_event_data)

    try:
        if len(financial_events_df) != 0:
            result = financial_events_df.pivot_table(index=['AmazonOrderId', 'SellerOrderId', 'MarketplaceName', 'PostedDate',
                                                            'SellerSKU', 'OrderItemId'], values=['QuantityShipped', 'Principal', 'Tax', 'GiftWrap', 'GiftWrapTax', 'GiftwrapChargeback',
                                                                                                 'ShippingCharge', 'ShippingTax', 'FBAPerUnitFulfillmentFee', 'Commission', 
                                                                                                 'ShippingChargeback'], aggfunc='sum').reset_index()
            return result
    except:
        if len(financial_events_df) != 0:
            result = financial_events_df.pivot_table(index=['AmazonOrderId', 'SellerOrderId', 'MarketplaceName', 'PostedDate',
                                                            'SellerSKU', 'OrderItemId'], values=['QuantityShipped', 'Principal', 'Tax', 'GiftWrap', 'GiftWrapTax',
                                                                                                 'ShippingCharge', 'ShippingTax', 'FBAPerUnitFulfillmentFee', 'Commission', 
                                                                                                 'ShippingChargeback'], aggfunc='sum').reset_index()
            return result
        
    else:
        return None
    

#Europe (Spain, UK, France, Belgium, Netherlands, Germany, Italy, Sweden, South Africa,
        #Poland, Saudi Arabia, Egypt, Turkey, United Arab Emirates, and India marketplaces)
url_europe = "https://sellingpartnerapi-eu.amazon.com/orders/v0/orders/{}/orderItems"
url_europe1 = "https://sellingpartnerapi-eu.amazon.com/orders/v0/orders/{}/address" #You can add your own url here if you want to fetch more data from the orders api

def complete_order(orders_df, refresh_token, base_url, base_url1, order_ids, batch_size):

    credentials = dict(
        refresh_token= refresh_token,  # From Seller central under Authorise -> Refresh Token
        lwa_app_id= client_id,  # From Seller Central, named CLIENT IDENTIFIER on website.
        lwa_client_secret= client_secret,  # From Seller Central, named CLIENT SECRET on website.
    )
    
    # Function to get access token
    def get_access_token(credentials):
        url = "https://api.amazon.com/auth/o2/token"
        payload = {
            "grant_type": "refresh_token",
            "client_id": credentials["lwa_app_id"],
            "client_secret": credentials["lwa_client_secret"],
            "refresh_token": credentials["refresh_token"]
        }
        headers = {
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"
        }
        response = requests.post(url, data=payload, headers=headers)
        response.raise_for_status()
        return response.json()["access_token"]
    
    # Get access token
    access_token = get_access_token(credentials)
    
    
    
    def fetch_order_list_items(order_ids_batch, access_token):
        #base_url = "https://sellingpartnerapi-fe.amazon.com/orders/v0/orders/{}/orderItems"
        headers = {
            "x-amz-access-token": access_token,
            "Content-Type": "application/json",
            "User-Agent": "YourAppName/1.0 (Language=Python/3.8)"
        }
        order_items_data = []
    
        for order_id in order_ids_batch:
            retries = 0
            while retries < 2:  # Retry at most once
                endpoint = base_url.format(order_id)
                response = requests.get(endpoint, headers=headers)
    
                if response.status_code == 200:
                    payload = response.json().get('payload', {})
                    order_items = payload.get('OrderItems', [])
                    
                    for item in order_items:
                        # Add the order_id to each item's dictionary
                        item['order_id'] = order_id
                        order_items_data.append(item)  # Collecting each item in the list
                    break  # Break the retry loop on success
                
                else:
                    retries += 1
                    if retries == 1:
                        print(f"Retrying for AmazonOrderId {order_id} after first failure...")
                        time.sleep(10)  # Wait for 10 seconds before retrying
                    else:
                        print(f"Failed to fetch orderListItems for AmazonOrderId {order_id} after retrying: {response.status_code}")
                        print(response.json())
                        break
            
            # Add a 0.5-second delay after each order ID request
            time.sleep(0.5)
        
        return order_items_data
    
    def fetch_order_list_items_in_batches(order_ids, batch_size, access_token):
        order_items_data = []
        num_batches = (len(order_ids) // batch_size) + 1
    
        for i in range(num_batches):
            start_idx = i * batch_size
            end_idx = min((i + 1) * batch_size, len(order_ids))
            order_ids_batch = order_ids[start_idx:end_idx]
    
            batch_order_items = fetch_order_list_items(order_ids_batch, access_token)
            order_items_data.extend(batch_order_items)
            
            print(f"Batch {i} processed.")
            
            # Add a delay of 10 seconds after each batch
            if i < num_batches - 1:  # No delay after the last batch
                time.sleep(10)
            
        return order_items_data
    
    # Collect all AmazonOrderId's into a list
    #order_ids = orders_df['AmazonOrderId'].unique().tolist()
    #batch_size = 20
    order_items = fetch_order_list_items_in_batches(order_ids, batch_size, access_token)
    print('Completed batch process')
    
    # Flatten the nested dictionaries
    def flatten_order_item(item):
        flattened_item = {}
        for key, value in item.items():
            if isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    flattened_item[f"{key}.{sub_key}"] = sub_value
            elif isinstance(value, list):
                flattened_item[key] = ', '.join(value)
            else:
                flattened_item[key] = value
        return flattened_item
    
    flattened_order_items = [flatten_order_item(item) for item in order_items]
    
    # Convert to DataFrame
    order_items_df = pd.DataFrame(flattened_order_items)
    # Replace NaN with 0
    order_items_df.fillna(0, inplace=True)
    
    def fetch_order_address(order_id, access_token):
        #base_url = "https://sellingpartnerapi-fe.amazon.com/orders/v0/orders/{}/address"
        headers = {
            "x-amz-access-token": access_token,
            "Content-Type": "application/json",
            "User-Agent": "YourAppName/1.0 (Language=Python/3.8)"
        }
        retries = 0
        while retries < 2:  # Retry at most once
            endpoint = base_url1.format(order_id)
            response = requests.get(endpoint, headers=headers)
    
            if response.status_code == 200:
                payload = response.json().get('payload', {})
                return {
                    "order_id": order_id,
                    "ShippingAddress": payload.get('ShippingAddress', {})
                }
            else:
                retries += 1
                if retries == 1:
                    print(f"Retrying for AmazonOrderId {order_id} after first failure...")
                    time.sleep(10)  # Wait for 10 seconds before retrying
                else:
                    print(f"Failed to fetch order address for AmazonOrderId {order_id} after retrying: {response.status_code}")
                    print(response.json())
                    return {"order_id": order_id, "ShippingAddress": {}}
            # Add a 0.3-second delay after each order ID request
            time.sleep(0.3)
    
    
    def fetch_order_addresses_in_batches(order_ids, batch_size, access_token):
        order_addresses_data = []
        num_batches = (len(order_ids) // batch_size) + 1
    
        for i in range(num_batches):
            start_idx = i * batch_size
            end_idx = min((i + 1) * batch_size, len(order_ids))
            order_ids_batch = order_ids[start_idx:end_idx]
    
            for order_id in order_ids_batch:
                order_address = fetch_order_address(order_id, access_token)
                order_addresses_data.append(order_address)
            
            print(f"Batch {i} processed.")
            
            # Add a delay of 10 seconds after each batch
            if i < num_batches - 1:  # No delay after the last batch
                time.sleep(10)
            
        return order_addresses_data
    
    # Collect all AmazonOrderId's into a list
    #order_ids = orders_df['AmazonOrderId'].unique().tolist()
    #batch_size = 20
    # Fetch order addresses in batches
    order_addresses = fetch_order_addresses_in_batches(order_ids, batch_size, access_token)
    print('Completed Batch1 Process')

    
    # Flatten the nested dictionaries
    def flatten_order_item(item):
        flattened_item = {}
        for key, value in item.items():
            if isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    flattened_item[f"{key}.{sub_key}"] = sub_value
            elif isinstance(value, list):
                flattened_item[key] = ', '.join(value)
            else:
                flattened_item[key] = value
        return flattened_item
    
    flattened_order_address = [flatten_order_item(item) for item in order_addresses]
    
    # Convert to DataFrame
    order_addresses_df = pd.DataFrame(flattened_order_address)
    # Replace NaN with 0
    order_addresses_df.fillna(0, inplace=True)
    
    #merging orderdetail with orderaddress
    merged_df = pd.merge(order_items_df, order_addresses_df, left_on='order_id', right_on='order_id', how='left')
    
    #merging finance with orderdetail and orderaddress
    final_sheet = pd.merge(
        orders_df, 
        merged_df, 
        left_on=['AmazonOrderId', 'SellerSKU'], 
        right_on=['order_id', 'SellerSKU'], 
        how='left'
    )

    return final_sheet


southafrica_data = finance_api(refresh_token, usa_data_start_time, usa_data_end_time, marketplace_southafrica)
time.sleep(2)

## STARTING BATCH PROCESS FOR ALL COUNTRIES.
#South-Africa
if southafrica_data is not None and not southafrica_data.empty:
    if 'AmazonOrderId' in southafrica_data.columns:
        southafrica_order_ids = southafrica_data['AmazonOrderId'].unique().tolist()
        print('South-Africa Order-ID')
        southafrica_order = complete_order(southafrica_data, refresh_token, url_europe, url_europe1, southafrica_order_ids, 20)
        time.sleep(2)
    else:
        # If 'AmazonOrderId' is not in the columns, pass
        print("South-Africa data does not contain 'AmazonOrderId' column. Skipping.")
else:
    print("South-Africa data is None or empty. Skipping.")


def my_southafrica(df):
    if df.empty:
        print("No data available. Function will not run.")
        return None
        
    cleaned_df = df.drop_duplicates(subset=[col for col in df.columns if col != 'date/time'])

    cleaned_df = cleaned_df.rename(columns={
                'PostedDate': 'date/time',
                'AmazonOrderId': 'order id',
                'SellerSKU': 'sku',
                'Title': 'Description',
                'MarketplaceName': 'marketplace',
                'QuantityOrdered': 'quantity',
                'TaxCollection.Model': 'tax collection model',
                'ShippingAddress.StateOrRegion' : 'order state',
                'ShippingAddress.City' : 'order city',
                'ItemPrice.Amount': 'product sales',
                'Tax': 'sales tax collected',
                'ShippingCharge': 'shipping credits',
                'ShippingTax': 'shipping credits tax',
                'GiftWrap': 'gift wrap credits',
                'GiftWrapTax': 'giftwrap credits tax',
                'PromotionDiscount.Amount': 'promotional rebates',
                'PromotionDiscountTax.Amount': 'promotional rebates tax',
                'Commission': 'selling fees',
                'FBAPerUnitFulfillmentFee': 'fulfilment by amazon fees'})

    if 'fba fees' in cleaned_df.columns:
        column_order = ['date/time', 'order id', 'sku', 'Description', 'marketplace', 'quantity', 'order city', 'order state', 
                        'ShippingDiscount.Amount','product sales','shipping credits', 'gift wrap credits','promotional rebates', 
                        'selling fees', 'fba fees']
        cleaned_df = cleaned_df[column_order]
    else:
        column_order = ['date/time', 'order id', 'sku', 'Description', 'marketplace', 'quantity', 'order city', 'order state', 
                        'ShippingDiscount.Amount','product sales','shipping credits', 'gift wrap credits','promotional rebates', 
                        'selling fees']
        cleaned_df = cleaned_df[column_order]
    
    cleaned_df['ShippingDiscount.Amount'] = cleaned_df['ShippingDiscount.Amount'].astype('float64')
    cleaned_df['promotional rebates'] = cleaned_df['promotional rebates'].astype('float64')
    
    # Add shippingDiscountAmount to Promotional rebates
    cleaned_df['promotional rebates'] = cleaned_df['promotional rebates'] + cleaned_df['ShippingDiscount.Amount']
    
    # Change the column data type to numeric
    cleaned_df['product sales'] = pd.to_numeric(cleaned_df['product sales'], errors='coerce')
    
    # Columns to ensure negative values
    columns_to_negative = ['promotional rebates']
    # Modify the values to be negative
    for col in columns_to_negative:
        cleaned_df[col] = cleaned_df[col].apply(lambda x: -abs(x))
    
    #columns to sum 
    if 'fba fees' in cleaned_df.columns:
        columns_to_sum = ['product sales', 'shipping credits', 'gift wrap credits', 'promotional rebates', 'selling fees', 'fba fees'] 
        # Sum and new column Total/
        cleaned_df['Total'] = cleaned_df[columns_to_sum].sum(axis=1)
    else:
        columns_to_sum = ['product sales', 'shipping credits', 'gift wrap credits', 'promotional rebates', 'selling fees'] 
        # Sum and new column Total/
        cleaned_df['Total'] = cleaned_df[columns_to_sum].sum(axis=1)

    return cleaned_df

def upload_file_to_sharepoint(ctx, folder_in_sharepoint, file_content, file_name):
    target_folder = ctx.web.get_folder_by_server_relative_url(folder_in_sharepoint)
    ctx.load(target_folder)
    ctx.execute_query()

    target_file = target_folder.upload_file(file_name, file_content)
    ctx.execute_query()
    print(f'File {file_name} uploaded successfully to {folder_in_sharepoint}')

try:
    southafrica_order
except NameError:
    print("There is no report for this country")
else:
    southafrica_transactions = my_southafrica(southafrica_order)
    if southafrica_transactions is not None:
        # Convert the DataFrame to a file-like object
        output = BytesIO()
        southafrica_transactions.to_excel(output, index=False)
        output.seek(0)

        # SharePoint credentials
        sharepoint_base_url = 'https://herbion.sharepoint.com/sites/1CData'
        sharepoint_user = #Add Your Sharepoint Username Here.
        sharepoint_password = #Add Your Sharepoint Password Here.
        folder_in_sharepoint = '/sites/1CData/Shared Documents/Amazon/Data/SAF - South Africa/South-Africa Daily Uploader'
        current_date = datetime.now().strftime("%Y%m%d")
        file_name = f'AmazonSouthAfrica_{current_date}.xlsx'

        # SharePoint authentication and upload
        auth = AuthenticationContext(sharepoint_base_url)
        if auth.acquire_token_for_user(sharepoint_user, sharepoint_password):
            ctx = ClientContext(sharepoint_base_url, auth)
            upload_file_to_sharepoint(ctx, folder_in_sharepoint, output, file_name)
        else:
            print("Authentication failed")