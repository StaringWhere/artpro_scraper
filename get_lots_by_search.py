import requests
from bs4 import BeautifulSoup
import bs4
from openpyxl import Workbook
import json
import re
import time
from requests.exceptions import ProxyError


# ----- Parameters -----
# Filename
name = "CindySherman"
count = 100

# Search Parameters
params = {
    'keyword': 'cindy Sherman',
    'start': '20',
    'count': '20',
    'lot_sort_type': 'LOT_SORT_BY_DEAL_PRICE',
    'optional_sort_order': 'SORT_ORDER_DESCEND',
    'search_type': 'SearchV2_MultiLot'
}

headers = {
    'authority': 'artpro.com',
    'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="96", "Google Chrome";v="96"',
    'x-locale': 'EN',
    'sec-ch-ua-mobile': '?0',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36',
    'accept': 'application/json, text/plain, */*',
    'grpc-metadata-token': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJqdGkiOjE2Nzk4fQ.zDILtF5NANNMprLg1tDIB1dZOfN7SsQRr6J_81aqcRg',
    'grpc-metadata-device': 'language=EN',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-site': 'same-origin',
    'sec-fetch-mode': 'cors',
    'sec-fetch-dest': 'empty',
    'referer': 'https://artpro.com/search?s=cindy%20Sherman',
    'accept-language': 'zh-CN,zh;q=0.9,ja;q=0.8',
    'cookie': 'uuid=rBAA62Gx/ji7LHfMA5HlAg==; _ga=GA1.1.351477267.1639054906; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJqdGkiOjE2Nzk4fQ.zDILtF5NANNMprLg1tDIB1dZOfN7SsQRr6J_81aqcRg; locale=en; _ga_CQTZ1PYDQ7=GS1.1.1639057990.2.1.1639061252.0',
}

# Initiate an excel sheet
workbook = Workbook()
sheet = workbook.active

# Sheet head
sheet["B1"] = "Artwork Name"
sheet["C1"] = "Artist Name"
sheet["D1"] = "Creation Time"
sheet["E1"] = "Auction Start Year"
sheet["F1"] = "Auction Start Time"
sheet["G1"] = "Auction Location"
sheet["H1"] = "Price Dealed"
sheet["I1"] = "Price Dealed (USD)"


# URL of search API
url = "https://artpro.com/web/api/v3/search"

# Iteration of pages
lotIndex = 1
haveMore = True
page = 0
while haveMore:

    # Get lots information JSON and decode
    params["count"] = count
    params["start"] = page * count
    response = requests.get(url, headers=headers, params=params)
    lotsString = response.text
    lotsJSON = json.loads(lotsString)

    haveMore = lotsJSON["data"]["have_more"]
    page += 1

    # Iteration of lots
    for lotJSON in lotsJSON["data"]["lots"]:

        artworkName = lotJSON["artwork_name"]
        artistNameEn = lotJSON["artist_name_en"]
        creationTimeStr = lotJSON["creation_time_str"]
        auctionStartTime = lotJSON["auction_start_time"]
        auctionStartTimeStr = time.strftime(r"%Y/%m/%d", time.localtime(int(auctionStartTime)))
        auctionStartYearStr = time.strftime(r"%Y", time.localtime(int(auctionStartTime)))
        auction_location = lotJSON["auction_location"]["city"]
        priceDealed = lotJSON["price_dealed"]
        priceDealedStr = priceDealed["currency"] + " " + str(priceDealed["amount"])
        priceDealedUsdStr = "USD" + " " + str(lotJSON["price_dealed_usd_amount"])

        # Write to excel
        sheet["A" + str(lotIndex + 1)] = lotIndex
        sheet["B" + str(lotIndex + 1)] = artworkName
        sheet["C" + str(lotIndex + 1)] = artistNameEn
        sheet["D" + str(lotIndex + 1)] = creationTimeStr
        sheet["E" + str(lotIndex + 1)] = auctionStartYearStr
        sheet["F" + str(lotIndex + 1)] = auctionStartTimeStr
        sheet["G" + str(lotIndex + 1)] = auction_location
        sheet["H" + str(lotIndex + 1)] = priceDealedStr
        sheet["I" + str(lotIndex + 1)] = priceDealedUsdStr

        # Print
        print(str(lotIndex) + " " + "completed")

        # End of each iteration
        lotIndex += 1


# Save excel
workbook.save(filename = name + str(int(time.time())) + ".xlsx")