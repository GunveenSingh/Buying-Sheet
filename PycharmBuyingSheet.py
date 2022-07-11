import pandas as pd

##Server computer loction is different
# df1 = pd.read_excel("F:\IB Australia Final Project\SalesData\SalesDataFiles\Sales_Last_6_Months.xlsx")
# df2 = pd.read_excel("F:\IB Australia Final Project\SalesData\SalesDataFiles\Historical_Data.xlsx")
try:
    df1=pd.read_excel("F:\PowerBI\IB Australia Final Project\SalesData\SalesDataFiles\Sales_Last_6_Months.xlsx")
except:
    df1=pd.read_excel("F:\IB Australia Final Project\SalesData\SalesDataFiles\Sales_Last_6_Months.xlsx")
try:
    df2=pd.read_excel("F:\PowerBI\IB Australia Final Project\SalesData\SalesDataFiles\Historical_Data.xlsx")
except:
    df2=pd.read_excel("F:\IB Australia Final Project\SalesData\SalesDataFiles\Historical_Data.xlsx")
    
df = df1.append(df2)
df = df[["Item", "Quantity", "Invoice Date"]]
df.head(10)

df.dtypes

import datetime

df['Invoice Date'] = pd.to_datetime(df['Invoice Date'])

#FilterDatesDF = pd.read_excel("BuyingSheetDates/StartEndDates.xlsx", sheet_name="Sheet1")
FilterDatesDF = pd.read_excel("P:\Buying Sheet Program\BuyingSheetDates/StartEndDates.xlsx", sheet_name="Sheet1")
FilterDatelist = (FilterDatesDF['StartDate YYYY-MM-DD'].tolist()) + (FilterDatesDF['EndDate YYYY-MM-DD'].tolist())

#start_date = "2018-01-01"
#end_date = "2021-08-31"
start_date = FilterDatelist[0]
end_date = FilterDatelist[1]

# df_filter=df[df["Invoice Date"]>=datetime.date(2020,1,1) and df["Invoice Date"]<=datetime.date(2020,12,31)]
df_filter = df[(df["Invoice Date"] >= start_date) & (df["Invoice Date"] <= end_date)]
df_filter.head()

df_filter['year'] = pd.DatetimeIndex(df_filter['Invoice Date']).year
df_filter['Invoice Date'] = pd.to_datetime(df_filter['Invoice Date'])
df_filter['month-year'] = df_filter['Invoice Date'].dt.strftime('%b-%Y')
df_filter['year_monthNo'] = pd.to_datetime(df_filter['Invoice Date']).dt.to_period('M')
df_filter = df_filter[["Item", "Quantity", "month-year", "year_monthNo"]]
df_filter.head(10)

buying_df = df_filter.groupby(["Item", "month-year", "year_monthNo"])["Quantity"].sum()
buying_df.head(10)

buying_df = buying_df.reset_index()
# del buying_df["index"]
buying_df.head(10)

# del buying_df["index"]
# buying_df.head(20)

buying_df = buying_df.sort_values(["year_monthNo", "Item"], ascending=(True, True))
buying_df = buying_df[["Item", "year_monthNo", "Quantity"]]
buying_df.head(10)

print("Preparing Daily Stock List.....")
import pandas as pd
import time
import csv
import glob
import os
import re
import io
from pandas import ExcelWriter
import numpy as np
import requests
import warnings

from pathlib import Path
import socket
from datetime import date

#####Fetching the stock after processing order
###Getting item details from MYOB
# Get UIDs of all items
# time.sleep(3)
# Get all item UID in ascending order

warnings.filterwarnings("ignore")

exception_list = pd.read_excel("P:/Stocklist/TheRejectList.xlsx", sheet_name="Sheet1")


# print(exception_list.head(10))


def hasCharacters(inputString):
    return bool(re.search(r'[a-zA-Z]', inputString))


url = "https://secure.myob.com/oauth2/v1/authorize/"

payload = 'client_id=tdd4dk347sw7nvdnr8tvbwta&client_secret=ZaCSqqus8DrEfC3fr3VvWr8C&grant_type=refresh_token&refresh_token=Pkv5%21IAAAAIKM9DYl342RjLtg3HZcfJn2zxEwrLBSUcjcm0ZRC98HsQAAAAHj3FUWHNzI_wS0ihflLHgVuDVqQNg7GNf4_TY3xUYRNN99_92IDs-h9IvwWiGO_lFPjc9HvWYlOUZtoiid1XC8u4VoZa3AruyasYVCVAhSh_qOi9qZ6zSQvPfqsOFEnFxPC5bQzHcMk7k1Ote7wxPzv8VxGKHHEw0C4sX8a0hnnQS7LGVgfTUvLIejFBQAjRIoD5sz_0lVZJoGnN3PPvU3rnVMICP3x-w7rQpXzoDi3w'
headers = {
    'Content-Type': 'application/x-www-form-urlencoded'
}

response = requests.request("POST", url, headers=headers, data=payload)

# print(response.text.encode('utf8'))
json_response = response.json()
# print(json_response)
token = json_response['access_token']
Authorization_replace = "Bearer " + token

url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number asc"
# test
# url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-9008') eq true&$orderby=Number asc"

payload = {}
headers = {
    'x-myobapi-key': 'dze2svn2nwsbrjzvmhwpsq57',
    'x-myobapi-version': 'v2',
    'Accept-Encoding': 'gzip,deflate',
    'Authorization': Authorization_replace
}

response = requests.request("GET", url, headers=headers, data=payload)
# print(response)
json_response = response.json()
# print(json_response)
item_UID = json_response['Items'][0]['Number']
# print(item_UID)
item_UID = json_response['Items']
# print(item_UID)
# item_df_aes = pd.DataFrame(columns=['UID','Item','unitWeight','BinLoc','SOH','baleQTY','Name'])
item_df_aes = pd.DataFrame(
    columns=['UID', 'Item No.', 'Item Name', 'Stock', 'Sold', 'Order', 'Avail', 'Bin Loc', 'Barcode', 'W/Price', 'Bale',
             'B/Price'])
for rows in item_UID:
    if(rows['IsActive'] == True):
        items_uid = rows['UID']
        items_number = rows['Number']
        items_name = rows['Name']
        # print('items_number',items_number)
        # print('CustomList3',rows['CustomList3'])
        # print('CustomField3',rows['CustomField3'])
        # Weight column not required
        #     if(rows['CustomList3']!=None):
        #         if(rows['CustomField3']!=None and hasCharacters(rows['CustomField3']['Value'])==False):
        #             if(float(rows['CustomField3']['Value'])!=0):
        #                 items_weight=float(rows['CustomList3']['Value'])/float(rows['CustomField3']['Value'])
        #             else:
        #                 items_weight=float(rows['CustomList3']['Value'])
        #         else:
        #             items_weight=float(rows['CustomList3']['Value'])
        #     else:
        #         items_weight=""
        # print('CustomList1',rows['CustomList1'])
        item_SOH = rows['QuantityOnHand']
        item_Sold = rows['QuantityCommitted']
        item_Order = rows['QuantityOnOrder']
        item_Avail = rows['QuantityAvailable']

        if (rows['CustomField1'] != None):
            items_location = rows['CustomField1']['Value']
        else:
            items_location = ""

        if (rows['CustomField2'] != None):
            items_barcode = rows['CustomField2']['Value']
        else:
            items_barcode = ""

        if (rows['CustomList2'] != None):
            # print(rows['CustomList2']['Value'])
            items_wholesale_price = float((re.findall('(0|\d{1,5}[.]?\d{1,5})', (rows['CustomList2']['Value'])))[0])
        else:
            items_wholesale_price = 0

        items_bale_price = round((float(items_wholesale_price) * 0.75), 2)

        # items_location=rows['CustomList1']['Value']
        # item_SOH=rows['QuantityAvailable']

        if (rows['CustomField3'] != None):
            items_bale_qty = rows['CustomField3']['Value']
        else:
            items_bale_qty = ""
        item_df_aes = item_df_aes.append(
            {'UID': items_uid, 'Item No.': items_number, 'Item Name': items_name, 'Stock': item_SOH, 'Sold': item_Sold,
             'Order': item_Order, 'Avail': item_Avail, 'Bin Loc': items_location, 'Barcode': items_barcode,
             'W/Price': items_wholesale_price, 'Bale': items_bale_qty, 'B/Price': items_bale_price}, ignore_index=True)
item_df_aes.head()
# print(item_df_aes[item_df_aes['Item No.']=='23-9008'].head())
item_df_aes.shape
# print(response.text.encode('utf8'))

# Get all item UID in descending order
time.sleep(1)
url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number desc"

payload = {}
headers = {
    'x-myobapi-key': 'dze2svn2nwsbrjzvmhwpsq57',
    'x-myobapi-version': 'v2',
    'Accept-Encoding': 'gzip,deflate',
    'Authorization': Authorization_replace
}

response = requests.request("GET", url, headers=headers, data=payload)
json_response = response.json()
# print(json_response)
item_UID = json_response['Items'][0]['Number']
# print(item_UID)
item_UID = json_response['Items']
# print(item_UID)
item_df_desc = pd.DataFrame(
    columns=['UID', 'Item No.', 'Item Name', 'Stock', 'Sold', 'Order', 'Avail', 'Bin Loc', 'Barcode', 'W/Price', 'Bale',
             'B/Price'])
for rows in item_UID:
    if (rows['IsActive'] == True):
        items_uid = rows['UID']
        items_number = rows['Number']
        items_name = rows['Name']
        # print('items_number',items_number)
        # print('CustomList3',rows['CustomList3'])
        # print('CustomField3',rows['CustomField3'])
        #     if(rows['CustomList3']!=None):
        #         if(rows['CustomField3']!=None and hasCharacters(rows['CustomField3']['Value'])==False):
        #             if(float(rows['CustomField3']['Value'])!=0):
        #                 items_weight=float(rows['CustomList3']['Value'])/float(rows['CustomField3']['Value'])
        #             else:
        #                 items_weight=float(rows['CustomList3']['Value'])
        #         else:
        #             items_weight=float(rows['CustomList3']['Value'])
        #     else:
        #         items_weight=""
        # print('CustomList1',rows['CustomList1'])
        item_SOH = rows['QuantityOnHand']
        item_Sold = rows['QuantityCommitted']
        item_Order = rows['QuantityOnOrder']
        item_Avail = rows['QuantityAvailable']

        if (rows['CustomField1'] != None):
            items_location = rows['CustomField1']['Value']
        else:
            items_location = ""

        if (rows['CustomField2'] != None):
            items_barcode = rows['CustomField2']['Value']
        else:
            items_barcode = ""

        if (rows['CustomList2'] != None):
            items_wholesale_price = float((re.findall('(0|\d{1,5}[.]?\d{1,5})', (rows['CustomList2']['Value'])))[0])
        else:
            items_wholesale_price = 0

        items_bale_price = round((float(items_wholesale_price) * 0.75), 2)

        # items_location=rows['CustomList1']['Value']
        # item_SOH=rows['QuantityAvailable']

        if (rows['CustomField3'] != None):
            items_bale_qty = rows['CustomField3']['Value']
        else:
            items_bale_qty = ""

        item_df_desc = item_df_desc.append(
            {'UID': items_uid, 'Item No.': items_number, 'Item Name': items_name, 'Stock': item_SOH, 'Sold': item_Sold,
             'Order': item_Order, 'Avail': item_Avail, 'Bin Loc': items_location, 'Barcode': items_barcode,
             'W/Price': items_wholesale_price, 'Bale': items_bale_qty, 'B/Price': items_bale_price}, ignore_index=True)
# item_df_desc.head()
item_df_desc.shape
# print(response.text.encode('utf8'))
##This dataframe contains the latest data
item_df = pd.merge(item_df_aes, item_df_desc,
                   on=['UID', 'Item No.', 'Item Name', 'Stock', 'Sold', 'Order', 'Avail', 'Bin Loc', 'Barcode',
                       'W/Price', 'Bale', 'B/Price'], how='outer')
# Matching with old dataframe

item_df = item_df[
    ['Item No.', 'Item Name', 'Stock', 'Sold', 'Order', 'Avail', 'Bin Loc', 'Barcode', 'W/Price', 'Bale', 'B/Price']]

item_df['W/Price'] = '$' + item_df['W/Price'].astype(str)
item_df['B/Price'] = '$' + item_df['B/Price'].astype(str)

#item_df['Item No.'] = item_df['Item No.'].str.upper()
item_df = item_df[~item_df['Item No.'].astype(str).str.startswith('23-X')]
# item_df = item_df[item_df['Barcode']!=""]
# item_df = item_df[item_df['Bin Loc']!=""]
item_df['Bin Loc'] = item_df['Bin Loc'].astype(str)
#item_df['Bin Loc'] = item_df['Bin Loc'].str.upper()
# item_df = item_df[item_df['Stock']!=0 and item_df['Sold']!=0 and item_df['Order']!=0 and item_df['Avail']!=0]
# item_df = item_df.loc[(item_df['Stock']!=0) & (item_df['Sold']!=0) & (item_df['Order']!=0) & (item_df['Avail']!=0)]
# item_df["stock_sum"]=item_df['Stock']+item_df['Sold']+item_df['Order']+item_df['Avail']
# item_df = item_df.loc[(item_df['stock_sum']!=0) & (item_df['Bin Loc']!="O/S")]
# del item_df['stock_sum']
item_df = item_df.sort_values(by=['Item No.'])
item_df.loc[
    (item_df['Item No.'].str.contains("PACK P|PACK S")) | (item_df['Item Name'].str.contains("Stand")), 'B/Price'] = \
    item_df['W/Price']

item_df.head(10)

temp_df = pd.merge(buying_df, item_df[["Item No.", "Item Name", "Stock"]], left_on=['Item'], right_on=['Item No.'],
                   how='outer')
temp_df['Item']= np.where(temp_df['Item'].isnull(), temp_df['Item No.'], temp_df['Item'])

del temp_df["Item No."]
buying_df = temp_df[["Item", "Item Name", "Stock", "year_monthNo", "Quantity"]]
buying_df.head(10)

buying_df = buying_df[buying_df['Item Name'].notna()]

# GET ORDERS

# purchaseOrderDF = pd.read_excel("PurchaseOrders/PurchaseOrderNumbers.xlsx", sheet_name="Sheet1")
purchaseOrderDF = pd.read_excel("P:/Buying Sheet Program/PurchaseOrders/PurchaseOrderNumbers.xlsx", sheet_name="Sheet1")

purchaseOrderDF.head(10)

filter_string = ""
purchaseOrderlist = purchaseOrderDF['PurchaseOrderNumber'].tolist()
for orders in range(len(purchaseOrderlist)):
    if (orders != len(purchaseOrderlist) - 1):
        filter_string = filter_string + "Number eq '" + str(purchaseOrderlist[orders]) + "' " + "or "
    else:
        filter_string = filter_string + "Number eq '" + str(purchaseOrderlist[orders]) + "'"
filter_string

import requests

url = "https://secure.myob.com/oauth2/v1/authorize/"

payload = 'client_id=tdd4dk347sw7nvdnr8tvbwta&client_secret=ZaCSqqus8DrEfC3fr3VvWr8C&grant_type=refresh_token&refresh_token=Pkv5%21IAAAAIKM9DYl342RjLtg3HZcfJn2zxEwrLBSUcjcm0ZRC98HsQAAAAHj3FUWHNzI_wS0ihflLHgVuDVqQNg7GNf4_TY3xUYRNN99_92IDs-h9IvwWiGO_lFPjc9HvWYlOUZtoiid1XC8u4VoZa3AruyasYVCVAhSh_qOi9qZ6zSQvPfqsOFEnFxPC5bQzHcMk7k1Ote7wxPzv8VxGKHHEw0C4sX8a0hnnQS7LGVgfTUvLIejFBQAjRIoD5sz_0lVZJoGnN3PPvU3rnVMICP3x-w7rQpXzoDi3w'
headers = {
    'Content-Type': 'application/x-www-form-urlencoded'
}

response = requests.request("POST", url, headers=headers, data=payload)

# print(response.text.encode('utf8'))
json_response = response.json()
# print(json_response)
token = json_response['access_token']
Authorization_replace = "Bearer " + token

# url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Purchase/Order/Item"

# url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Purchase/Order/Item/?$filter=(Number eq 'FW/2266' or Number eq 'WM23IB07' or Number eq 'FW/2305')"
# url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Sale/Order/Item/?$filter=Number eq"+"'"+"0"+user_invoice+"'
url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Purchase/Order/Item/?$filter=(" + filter_string + ")"

# ?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number asc"

payload = {}
headers = {
    'x-myobapi-key': 'tdd4dk347sw7nvdnr8tvbwta',
    'x-myobapi-version': 'v2',
    'Accept-Encoding': 'gzip,deflate',
    'Authorization': Authorization_replace
}

response = requests.request("GET", url, headers=headers, data=payload)

# print(response.text)

json_response = response.json()

print(json_response)

item_PO = json_response['Items']

print(len(item_PO))

item_PO = json_response['Items']
PO_temp_df = buying_df
#for orders in range(len(purchaseOrderlist)):

for order in item_PO:

    PO_name=str(order['Number'])+"\n Ordr Dt:"+str(order['Comment'])+"\n ETA:"+str(order['PromisedDate'])+"\n ETD:"+str(order['JournalMemo'])+"\n Cont. Size:"+str(order['ShippingMethod'])
    #PO_df = pd.DataFrame(columns=['Item', order['Number']])
    PO_df = pd.DataFrame(columns=['Item', PO_name])
    for line in order['Lines']:
        if (line['Type'] == 'Transaction'):
            item_number = line['Item']['Number']
            item_quantity = line['BillQuantity']
            PO_df = PO_df.append({'Item': item_number, PO_name: item_quantity}, ignore_index=True)

    PO_temp_df = pd.merge(PO_temp_df, PO_df, on=['Item'], how='left')
    # PO_df.head()

print(PO_temp_df.head(10))

buying_df = PO_temp_df
# order_col_names = ""
# for orders in range(len(purchaseOrderlist)):
#     if (orders != len(purchaseOrderlist) - 1):
#         order_col_names = order_col_names + "'" + str(purchaseOrderlist[orders]) + "'" + ","
#     else:
#         order_col_names = order_col_names + "'" + str(purchaseOrderlist[orders]) + "'"
# all_col_names=
# columns_list = []
# columns_list = ['Item', 'Item Name', 'Stock']
# for items in purchaseOrderlist:
#     columns_list.append(items)
# index_columns_list = columns_list
#
# print(index_columns_list)
#
# all_columns_list = columns_list
# all_columns_list.append('year_monthNo')
#all_columns_list.append('Quantity')
#
# print(all_columns_list)

cols_at_end = ['year_monthNo', 'Quantity']
buying_df = buying_df[[c for c in buying_df if c not in cols_at_end] + [c for c in cols_at_end if c in buying_df]]

#buying_df = buying_df[all_columns_list]

buying_df['year_monthNo']= np.where(buying_df['year_monthNo'].isnull(), "1900-01", buying_df['year_monthNo'])

buying_df.head(10)

buying_df.dtypes
index_columns_list=list(buying_df)
index_columns_list.remove('Quantity')

index_columns_list.remove('year_monthNo')

#print(index_columns_list)

buying_df = buying_df.fillna(0)
buying_df.head(10)

buying_df_pivot = pd.pivot_table(buying_df, values='Quantity', index=index_columns_list,
                                 columns='year_monthNo').reset_index()

# buying_df_pivot = buying_df_pivot[sorted(buying_df_pivot)]
buying_df_pivot.head()

buying_df_pivot = buying_df_pivot.fillna(0)
# del buying_df_pivot['year_monthNo']
buying_df_pivot.head()

buying_df_pivot.tail()

search = "23-"
# boolean series returned

buying_df_pivot = buying_df_pivot[buying_df_pivot["Item"].str.startswith(search)]
buying_df_pivot.head(10)

prefixes = ['23-X', '23-x', 'x']

buying_df_pivot = buying_df_pivot[~buying_df_pivot.Item.str.startswith(tuple(prefixes))]
buying_df_pivot.head(10)

buying_df_pivot.tail(10)

#start_position = len(purchaseOrderDF) + 4 - 1
# Updating logic for number of POs
start_position = len(item_PO) + 4 - 1
start_position

buying_df_pivot['SALES'] = buying_df_pivot.iloc[:, start_position:].sum(axis=1)

buying_df_pivot['SALES'] = buying_df_pivot['SALES'].astype("int")

buying_df_pivot['No. of Months'] = buying_df_pivot.iloc[:, start_position:-1].select_dtypes(np.number).gt(0).sum(axis=1)

buying_df_pivot['AVG PM'] = buying_df_pivot['SALES'] / buying_df_pivot['No. of Months']
#Handle division by zero
buying_df_pivot['AVG PM'] = buying_df_pivot['AVG PM'].apply(lambda x: 0 if x == np.inf else x)

buying_df_pivot['AVG PM'] = buying_df_pivot['AVG PM'].fillna(0)

buying_df_pivot['AVG PM'] = buying_df_pivot['AVG PM'].round(0).astype("int")

buying_df_pivot['AVG PM'].head(10)

buying_df_pivot['Months Supply'] = buying_df_pivot['Stock'] / buying_df_pivot['AVG PM']
#Handle division by zero
buying_df_pivot['Months Supply'] = buying_df_pivot['Months Supply'].apply(lambda x: 0 if x == np.inf else x)

buying_df_pivot['Months Supply'] = buying_df_pivot['Months Supply'].round(0)
buying_df_pivot['Months Supply'] = buying_df_pivot['Months Supply'].fillna(0)

buying_df_pivot['3 Months Sales Re-Ord'] = buying_df_pivot['AVG PM'] * 3

buying_df_pivot['3 Months Sales Re-Ord'].head(10)

timestr = time.strftime("%d%m%Y")
buyingSheet_headder = "I.B. Australia Buying Sheet - " + timestr
filename = "BuyingSheet_" + timestr + ".xlsx"
writer = pd.ExcelWriter(filename, engine='xlsxwriter')
# buying_df_pivot.to_excel(writer, sheet_name='Sheet1',index=False,header=None)
buying_df_pivot.to_excel(writer, sheet_name='Sheet1', index=False)
workbook = writer.book
worksheet = writer.sheets['Sheet1']

writer.save()

#quotes_df = pd.read_excel("Quotes/Quotes_1120-0821_full.xlsx")
quotes_df = pd.read_excel("P:/Buying Sheet Program/Quotes/Quotes.xlsx")



print(quotes_df.head(10))

quotes_df_cleaned = quotes_df.iloc[10:, 0:11]

quotes_df_cleaned.columns = ['Item', 'Supplier', 'Purchase No.', 'Date', 'Quantity', 'Status', 'Promised Date', 'Supplier Inv No.', 'Memo', 'Comments','Ship Via']

quotes_df_cleaned.loc[quotes_df_cleaned['Quantity'].isnull(), 'Item'] = quotes_df_cleaned['Supplier']

quotes_df_cleaned['Item'] = quotes_df_cleaned['Item'].ffill(axis=0)

quotes_df_cleaned = quotes_df_cleaned[quotes_df_cleaned['Status'].notna()]

quotes_df_cleaned.head(10)

quotes_df_final = quotes_df_cleaned[['Item', 'Quantity', 'Supplier Inv No.','Comments','Promised Date','Memo','Ship Via']]

quotes_df_final['header'] = quotes_df_final['Supplier Inv No.'] + "\n Ordr Dt:" + quotes_df_final['Comments'] + "\n ETA:" + quotes_df_final['Promised Date'] + "\n ETD:" + quotes_df_final['Memo'] + "\n Cont. Size" + quotes_df_final['Ship Via']

quotes_df_final = quotes_df_final[['Item', 'Quantity', 'header']]

quotes_df_final.head(10)

quotes_df_final['Quantity'] = quotes_df_final['Quantity'].astype('int')

quotes_df_pivot = pd.pivot_table(quotes_df_final, values='Quantity', index=['Item'],
                                 columns='header').reset_index()
quotes_df_pivot = quotes_df_pivot.fillna(0)

print(quotes_df_pivot.head(10))

buying_sheet_full = pd.merge(buying_df_pivot, quotes_df_pivot, on=['Item'], how='left')

buying_sheet_full = buying_sheet_full.fillna(0)

print(buying_sheet_full.head(10))

del buying_sheet_full['1900-01']

timestr = time.strftime("%d%m%Y")
buyingSheet_headder = "I.B. Australia Buying Sheet - " + timestr
#filename = "BuyingSheet/BuyingSheet_" + timestr + ".xlsx"
#filename = "BuyingSheet/BuyingSheet" + ".xlsx"
filename = "P:\Buying Sheet Program\BuyingSheet\BuyingSheet" + ".xlsx"

writer = pd.ExcelWriter(filename, engine='xlsxwriter')
# buying_df_pivot.to_excel(writer, sheet_name='Sheet1',index=False,header=None)
buying_sheet_full.to_excel(writer, sheet_name='Sheet1', index=False)
workbook = writer.book
worksheet = writer.sheets['Sheet1']

writer.save()







