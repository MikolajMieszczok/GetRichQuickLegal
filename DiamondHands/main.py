from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter
import yfinance
from datetime import datetime, timedelta


#tu się źle zapisywało, brało też nazwy firm w które inwestowali, ale usunąłem
def filter_politicians(array):
    for x in range(len(array)):
        if(x%2 != 0):
            array[x]="0"

def write_data_into_column(worksheet_temp, arr, column):
    worksheet_temp.set_column(str(column) + ":" + str(column), 20)
    for i, data in enumerate(arr):
        worksheet_temp.write(str(column) + str(i), data)

def remove_even_indexes(arr):
    return [arr[i] for i in range(len(arr)) if i % 2 != 0]

def remove_odd_indexes(arr):
    return [arr[i] for i in range(len(arr)) if i % 2 == 0]

def convert_date_format(date_str):
    month_map = {
        "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04","May": "05", "Jun": "06", "Jul": "07", "Aug": "08", "Sept": "09", "Oct": "10", "Nov": "11", "Dec": "12"
    }
    day, month = date_str.split()
    return f"{day}.{month_map[month]}."

def merge_dates(date_str, date_filed_str):
    return str(date_filed_str) + str(date_str)

def remove_marital_status(year_bought):
    indices_to_keep = [
        i for i in range(len(year_bought)) if i%3==0
    ]
    year_bought[:] = [year_bought[i] for i in indices_to_keep]

def remove_all_non_buy(politicians, stocks, filed_after, whether_buy, date_filled, year_bought):
    indices_to_keep = [
        i for i in range(len(whether_buy)) if whether_buy[i] != 0 and stocks[i] != 'N/A'
    ]

    # Update lists in place
    politicians[:] = [politicians[i] for i in indices_to_keep]
    stocks[:] = [stocks[i] for i in indices_to_keep]
    filed_after[:] = [filed_after[i] for i in indices_to_keep]
    whether_buy[:] = [whether_buy[i] for i in indices_to_keep]
    date_filled[:] = [date_filled[i] for i in indices_to_keep]
    year_bought[:] = [year_bought[i] for i in indices_to_keep]
# def change_date_format_and_read_date_of_posting(date_filed, filed_after):

def remove_items(test_list, item):
    # using list comprehension to perform the task
    res = [i for i in test_list if i != item]
    return res

url_start = "https://www.capitoltrades.com/trades?page="
politicians = []
stocks = []
politicians_idx = []
stocks_idx = []
filed_after = []
whether_buy = []
date_filed = []
amount_bought = []
year_bought = []
temp_date = []
temp_stock = []
todays_date = []
sixty_days_ago = []
columns = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO',
 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD',
 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ']

for x in range(1, 3):
    url = url_start + str(x)
    page = urlopen(url)
    html_bytes = page.read()
    html = html_bytes.decode("utf-8")
    soup = BeautifulSoup(html, 'html.parser')
    politician_links = soup.find_all('a', class_='text-txt-interactive')
    stock_links = soup.find_all('span', class_='q-field issuer-ticker')
    filed_after_links = soup.find_all('span', class_=['reporting-gap-tier--1', 'reporting-gap-tier--2', 'reporting-gap-tier--3', 'reporting-gap-tier--4'])
    whether_buy_links = soup.find_all('span', class_=['q-field tx-type tx-type--buy', 'q-field tx-type tx-type--sell', 'q-field tx-type tx-type--buy has-asterisk', 'q-field tx-type tx-type--sell has-asterisk', 'q-field tx-type tx-type--receive has-asterisk'])
    date_filed_links = soup.find_all('div', class_='text-size-3 font-medium')
    amount_bought_links = soup.find_all('span', class_='mt-1 text-size-2 text-txt-dimmer hover:text-foreground')
    year_bought_links = soup.find_all('div', class_='text-size-2 text-txt-dimmer')
    for link in politician_links:
        name = link.text
        politicians.append(name)
    for link in stock_links:
        name = link.text
        stocks.append(name)
    for link in filed_after_links:
        name = link.text
        filed_after.append(name)
    for link in whether_buy_links:
        name = link.text
        if name == "buy":
            whether_buy.append(1)
        else:
            whether_buy.append(0)
    for link in date_filed_links:
        name = link.text
        date_filed.append(name)
    for link in year_bought_links:
        name = link.text
        year_bought.append(name)

remove_marital_status(year_bought)
year_bought_temp = remove_items(year_bought, "Joint")
year_bought_temp = remove_items(year_bought_temp, "Spouse")
date_filed = remove_odd_indexes(date_filed)
filter_politicians(politicians)
res_temp = remove_items(politicians, 0)
res = remove_items(res_temp, "0")
remove_all_non_buy(res, stocks, filed_after, whether_buy, date_filed, year_bought)
print("dane po zmianie\n\n\n\n")

data_trans = []
for filed in range(len(temp_date)):
    data_trans.append(merge_dates(year_bought[filed], temp_date[filed]))
date_filed = data_trans
for stock in stocks:
    temp_stock.append(stock.replace(":", "."))
stocks = temp_stock
p = 0
p1 = 0
p2 = 0
p3 = 0
p4 = 0
p5 = 0
for politician in res:
    print(politician)
    p += 1
for stock in stocks:
    print(stock)
    p1 += 1
for filed in filed_after:
    print(filed)
    p2 += 1
for if_buy in whether_buy:
    print(if_buy)
    p3 += 1
for date in date_filed:
    print(date)
    p4 += 1
for year in year_bought:
    print(year)
    p5 += 1
print(p)
print(p1)
print(p2)
print(p3)
print(p4)
print(p5)


unique_ids_polit = {name: hash(name) for name in set(politicians)}
unique_ids_stock = {stock: hash(stock) for stock in set(stocks)}
print(unique_ids_polit)
print(unique_ids_stock)
for polit in res:
    politicians_idx.append(unique_ids_polit[polit])
for stock in stocks:
    stocks_idx.append(unique_ids_stock[stock])
print(politicians_idx)
print(stocks_idx)
for x in range(len(year_bought)):
    todays_date.append(datetime.now().strftime("%Y-%m-%d"))
for x in range(len(year_bought)):
    sixty_days_ago.append((datetime.now() + timedelta(days=-60)).strftime("%Y-%m-%d"))

#Transfering Data to EXCEL
workbook = xlsxwriter.Workbook("C:/Users/Mikołaj/Desktop/FIRSTEXCEL.xlsx")
worksheet = workbook.add_worksheet()
write_data_into_column(worksheet, res, "A")
write_data_into_column(worksheet, stocks, "B")
write_data_into_column(worksheet, filed_after, "C")
write_data_into_column(worksheet, todays_date, "D")
write_data_into_column(worksheet, sixty_days_ago, "E")
workbook.close()