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
# URL of lots information JSON (no matter which page)
url = "https://artpro.com/web/api/v3/get_lot_by_artist?artist_id=dAcYmTXWcyVTOEt85t7qh8rTw8SyMfY7fEbTGjFZPSObfjycdRfKPS6kt1&start=0&count=20&optional_sort_order=SORT_ORDER_DESCEND&count_only=false"
count = 100

headers = {
    "cookie": r"uuid=rBAA62Gx/ji7LHfMA5HlAg==; _ga=GA1.1.351477267.1639054906; token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJqdGkiOjE2Nzk4fQ.zDILtF5NANNMprLg1tDIB1dZOfN7SsQRr6J_81aqcRg; locale=en; _ga_CQTZ1PYDQ7=GS1.1.1639057990.2.0.1639057990.0",
    "grpc-metadata-device": "language=EN",
    "grpc-metadata-token": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJqdGkiOjE2Nzk4fQ.zDILtF5NANNMprLg1tDIB1dZOfN7SsQRr6J_81aqcRg"
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

# Iteration of pages
lotIndex = 1
haveMore = True
page = 0
while haveMore:

    # Get lots information JSON and decode
    response = requests.get(re.sub(r'(.*start=)[0-9]+(&count=)[0-9]+(.*)', r'\g<1>' + str(page * count) + r'\g<2>' + str(count) + r'\g<3>', url), headers = headers)
    lotsString = response.text
    lotsJSON = json.loads(lotsString)

    haveMore = lotsJSON["data"]["have_more"]
    page += 1

    # Iteration of lots
    for lotJSON in lotsJSON["data"]["list"]:

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