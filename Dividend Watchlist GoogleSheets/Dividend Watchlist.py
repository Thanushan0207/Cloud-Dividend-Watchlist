import yfinance as yf
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Authenticate with the Google Sheets API
scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name(r'D:\Dividend Watchlist GoogleSheets\service_account.json', scope)
client = gspread.authorize(creds)

# Open the target sheet and select the first worksheet
sheet = client.open('Dividend Watchlist').sheet1

# Define the tickers and the corresponding data keys
tickers = ['MSFT', 'AAPL', 'TSLA', 'ITE.TO', 'SUN']
data_keys = {
    'longName': 1,
    'symbol': 2,
    'sector': 4,
    'industry': 5,
    'currentPrice': 6,
    'financialCurrency': 7,
    'marketCap': 8,
    'dividendYield': 9,
    'dividendRate': 10,
    'payoutRatio': 11,
    'exDividendDate': 12,
    'lastDividendValue': 13,
    'lastDividendDate': 14
}

# Retrieve the data from Yahoo Finance and write it to the sheet
for i, ticker in enumerate(tickers):
    ticker_info = yf.Ticker(ticker).info
    for key, col in data_keys.items():
        if key in ticker_info:
            if key == 'exDividendDate' or key == 'lastDividendDate':
                date_str = ticker_info[key]
                if date_str is not None:
                    date = datetime.fromtimestamp(date_str).strftime('%Y-%m-%d')
                    sheet.update_cell(i+2, col, date)
                else:
                    sheet.update_cell(i+2, col, 'None')
            else:
                sheet.update_cell(i+2, col, ticker_info[key])
        else:
            sheet.update_cell(i+2, col, 'None')

print("Data written to sheet successfully.")