import requests
import xlwings as xw

OANDA_API_KEY = "7203735f4489a07b1fdaa82e0825b643-037a13c8ecd863561ee2fb38166bda77"
OANDA_URL = "https://api-fxpractice.oanda.com/v3/instruments/{}/candles"
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

@xw.sub  # This makes the function callable from Excel
def update_oanda_data():
    wb = xw.Book.caller()             # The workbook calling the macro
    sht = wb.sheets["Data"]           # Sheet where you want data
    records = [get_latest_candle(s) for s in symbols]

    # Find the first empty row
    last_row = sht.range("A" + str(sht.cells.last_cell.row)).end("up").row + 1

    # Append records below existing data
    sht.range(f"A{last_row}").value = records
