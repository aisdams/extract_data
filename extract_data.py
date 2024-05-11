import os
import requests
import openpyxl
from datetime import date

# Fungsi untuk mengambil data dari API
def get_api_data():
    url = 'https://anypoint.mulesoft.com/mocking/api/v1/sources/exchange/assets/1d520d19-e794-43c5-a112-7464af214520/exp-ecommerce/1.0.53/m/oauth/products'
    headers = {
        'sourceSystem': 'SITECORE',
        'countryCode': 'ANY',
        'brandCode': 'HAIBIKE',
        'channelReference': 'D2C',
        'customerId': 'A400312001'
    }
    response = requests.get(url, headers=headers)
    data = response.json()
    return data

# Fungsi untuk menulis data ke file Excel
def write_to_excel(data):
    today = date.today()
    count = 1
    while True:
        excel_file = f'data_{today}_{count}.xlsx'
        if not os.path.exists(excel_file):
            break
        count += 1
    
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['Product ID', 'Stock Status', 'Quantity', 'Price Amount', 'Price Currency', 'Remark'])
    for product in data:
        if 'stock' in product and 'details' in product['stock'] and product['stock']['details']:
            stock_details = product['stock']['details'][0]
            sheet.append([
                product['id'],
                stock_details['status'],
                stock_details['quantity'],
                product['price']['amount'],
                product['price']['currency'],
                product['remark']
            ])
    workbook.save(excel_file)

# Main program
if __name__ == "__main__":
    api_data = get_api_data()
    write_to_excel(api_data)