import datetime
import xlsxwriter
import yfinance as yf
import openpyxl
from datetime import datetime, timedelta
# def stock_price_over_time(row, ws, worksheet):
#     ticker = ws["C" + row].value.split(".")[0]
#     start_date = datetime.datetime.strptime(ws["F" + row].value, "%Y-%m-%d").date()
#     end_date = datetime.datetime.strptime(ws["H" + row].value, "%Y-%m-%d").date()
#     hist = yf.Ticker(ticker).history(start=start_date, end=end_date, interval="1d")
#     if not hist.empty:
#         last_close = hist["Close"].iloc[-1]  # Get the last available closing price
#     else:
#         last_close = "N/A"  # Or handle it as needed
#     worksheet.write("I" + row, last_close)
columns = ['H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO',
 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD',
 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ']
workbook = xlsxwriter.Workbook('PRAWIEGOTOWE.xlsx')
worksheet = workbook.add_worksheet()
wb = openpyxl.load_workbook("STOCKS-DATE.xlsx")
ws = wb.active
for row in range(12000, 14507):
    try:
        print(row)
        row = str(row)
        not_filtered = yf.Ticker(str(ws["B" + row].value).split(".")[0]).history(start = str(((ws["D" + row]).value)).split(" ")[0], end = str(((ws["C" + row]).value)).split(" ")[0], interval = "1d")
        #print(not_filtered)
        if not_filtered.empty:
            continue
        filtered = not_filtered["Open"]
        for col, price in enumerate(filtered):
            worksheet.write(columns[col] + row, price)
    except Exception as e:
        print(f"Error at row {row}: {e}")
workbook.close()
#15238
