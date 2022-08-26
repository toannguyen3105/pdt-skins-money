from openpyxl import Workbook
import csv
import os.path
import scrapy
import glob
from scrapy.http import Request

# Min price is 10$
MIN_PRICE = 1000

# Max price is 30000$
MAX_PRICE = 3000000

# LIMIT
LIMIT = 60

# HEADERS
headers = {
    'authority': 'skinsmonkey.com',
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'en-US,en;q=0.9,vi;q=0.8',
    'cookie': 'sm_locale=en; sm_id=juU2jTgbwO2JCIXAx6KBVKV4iYCf2Q0W; _ga=GA1.1.295454533.1661012906; FPID=FPID2.2.Tfka79m9BUgYA4Dhgytr1J5o0hSw2wH9h60ZLLtJYkY%3D.1661012906; _hjSessionUser_2621449=eyJpZCI6IjI1YTQxMGRmLWVjMjYtNWZjZS1iMDZkLTJkODM0YWQyZmM1NiIsImNyZWF0ZWQiOjE2NjEwMTI5MDU0NTUsImV4aXN0aW5nIjp0cnVlfQ==; _hjIncludedInSessionSample=0; _hjSession_2621449=eyJpZCI6IjI0ZGE0MTE3LTYwOGQtNGE4Zi1hOTc1LWFlNzc3ZDVjZjI5ZiIsImNyZWF0ZWQiOjE2NjEzNjA5Nzk5ODAsImluU2FtcGxlIjpmYWxzZX0=; _hjAbsoluteSessionInProgress=1; FPLC=Asnd1PCY%2BkUHTDP7h5%2B6UhviCT935nT3YY2pYLZR2YQnVUTVcojCTiI3P52wUYGybLjAHzMGIebAVVx8joU%2F9GAAgP%2BxfpQi9PCDupNwdheSI%2FIaiPKQNqZrwt64GA%3D%3D; crisp-client%2Fsession%2Fc65b3c05-d915-4ebe-a091-3a9d4d62bcb5=session_1d75a4ba-47e3-4d8c-bfab-059ec08e6819; _ga_YF6RLKGLVH=GS1.1.1661363015.4.1.1661363017.0.0.0',
    'referer': 'https://skinsmonkey.com/trade?fbclid=IwAR05YQHObBSKJBrEUmFVUhO7in8VdK8kh8COBpXZAGK2G1_Om9MBrFZYiyo',
    'sec-ch-ua': '".Not/A)Brand";v="99", "Google Chrome";v="103", "Chromium";v="103"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Linux"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36'
}


class ItemsSpider(scrapy.Spider):
    name = 'items'
    allowed_domains = ['skinsmonkey.com']

    # OFFSET
    OFFSET = 0

    def start_requests(self):
        yield Request(
            url='https://skinsmonkey.com/api/inventory?limit='
                + str(LIMIT) + '&offset=0&appId=730&virtual=false&sort=price-desc&priceMin='
                + str(MIN_PRICE) + '&priceMax=' + str(MAX_PRICE) + '&featured=false',
            headers=headers
        )

    def parse(self, response):
        result = response.json()
        items = result['assets']
        if len(items) > 0:
            for item in items:
                title = item['item']['marketName']
                price = item['item']['price']

                yield {
                    'name': title,
                    'price': price,
                }

            for number in range(10000):
                yield Request(
                    response.urljoin(
                        'inventory?limit={}&offset={}&appId=730&virtual=false&sort=price-desc&priceMin={}&priceMax={}&featured=false'.format(
                            LIMIT, number * LIMIT, MIN_PRICE, MAX_PRICE)),
                    headers=headers,
                    callback=self.parse,
                )

    def close(self, reason):
        csv_file = max(glob.iglob('*csv'), key=os.path.getctime)

        wb = Workbook()
        ws = wb.active

        with open(csv_file, 'r', encoding='utf-8') as f:
            for row in csv.reader(f):
                ws.append(row)

        wb.save(csv_file.replace('.csv', '') + '.xlsx')
