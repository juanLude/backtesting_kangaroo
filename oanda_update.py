import requests
import os
from dotenv import load_dotenv
# import xlwings as xw

# load .env in development
load_dotenv()  # safe: if .env missing, it just uses existing environment vars

OANDA_API_KEY = os.environ["OANDA_API_KEY"]
OANDA_URL = os.environ.get(
    "OANDA_URL",
    "https://api-fxpractice.oanda.com/v3/instruments/{}/candles"
)
symbols = ["XAU_USD", "AUD_USD"]

def get_latest_candle(symbol):
    headers = {"Authorization": f"Bearer {OANDA_API_KEY}"}
    params = {"count": 1, "granularity": "D", "price": "M"}
    response = requests.get(OANDA_URL.format(symbol), headers=headers, params=params)
    candle = response.json()["candles"][0]
    return [symbol, candle["time"],
            float(candle["mid"]["o"]),
            float(candle["mid"]["h"]),
            float(candle["mid"]["l"]),
            float(candle["mid"]["c"])]

# @xw.sub  # This makes the function callable from Excel
# def update_oanda_data():
#     wb = xw.Book.caller()             # The workbook calling the macro
#     sht = wb.sheets["Data"]           # Sheet where you want data
#     records = [get_latest_candle(s) for s in symbols]

#     # Find the first empty row
#     last_row = sht.range("A" + str(sht.cells.last_cell.row)).end("up").row + 1

#     # Append records below existing data
#     sht.range(f"A{last_row}").value = records


# call the function directly for testing and showcase
if __name__ == "__main__":
    for symbol in symbols:
        print(get_latest_candle(symbol))
    # convert to DataFrame for exploration
    import pandas as pd
    df = pd.DataFrame([get_latest_candle(symbol) for symbol in symbols])
    print(df)