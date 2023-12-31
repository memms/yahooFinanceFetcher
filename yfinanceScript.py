import yfinance as yf
import xlsxwriter as xl
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta, FR, MO, SA
import pandas as pd

# Script made to get stock data from yahoo finance and export it to excel
# Get Friday's price, 5 day high, 5 day low, 3 month high, 3 month low
# Calculate % Gain/Loss given purchase price compared to Friday's price
# Excel sheet Colums: Stock Ticker, Purchase Price, Friday's Price, 5 day high, 5 day low, 3 month high, 3 month low, % Gain/Loss

# Change these values
# Format: 'Ticker': purchase price
# Seperate by comma (',')
tickersDict = {'AAPL':100, 'NFLX': 450}
# This determines the threshold for the % Gain/Loss to be highlighted in red or green in the excel sheet
gainLossMargin = 5


# -------- DO NOT CHANGE BELOW THIS LINE-------------------
class Stock:
    def __init__(self, ticker, purchase_price, friday_price,
                 wk_high, wk_low, mo_high,
                 mo_low, percent_gain_loss):
        self.ticker = ticker
        self.purchase_price = purchase_price
        self.friday_price = friday_price
        self.percent_gain_loss = percent_gain_loss
        self.wk_high = wk_high
        self.wk_low = mo_low
        self.mo_high = mo_high
        self.mo_low = wk_low

final = []

def get_stock_data(tickers, start, end):
    return yf.download(tickers, start = start, end = end, interval='1d')


def get_highs_lows(stock: list):
    return stock['High'].max(), stock['Low'].min()

def get_percent_gain_loss(purchase_price: list, friday_price: list):
    return (friday_price - purchase_price) / purchase_price * 100


def create(tickers: dict, firday_price, wk_high, wk_low, mo_high, mo_low, percent_gain_loss):
    for key, value in tickers.items():
        final.append(Stock(key, value, firday_price[key],
                           wk_high[key], wk_low[key], mo_high[key], mo_low[key], percent_gain_loss[key]))
        
def export_to_excel():
    xls = xl.Workbook('stockData.xlsx')
    sheet = xls.add_worksheet()
    cell_format_red = xls.add_format({'font_color': 'red'})
    cell_format_green = xls.add_format({'font_color': 'green'})
    # Make headers
    sheet.write(0, 0, 'Ticker')
    sheet.write(0, 1, 'Purchase')
    sheet.write(0, 2, 'Friday')
    sheet.write(0, 3, 'Week High')
    sheet.write(0, 4, 'Week Low')
    sheet.write(0, 5, '3 Month High')
    sheet.write(0, 6, '3 Month Low')
    sheet.write(0, 7, '% Gain/Loss')
    
    
    # Write data
    for i in range(len(final)):
        ind = i + 1
        sheet.write(ind, 0, final[i].ticker)
        sheet.write(ind, 1, final[i].purchase_price)
        sheet.write(ind, 2, final[i].friday_price.round(2))
        sheet.write(ind, 3, final[i].wk_high.round(2))
        sheet.write(ind, 4, final[i].wk_low.round(2))
        sheet.write(ind, 5, final[i].mo_high.round(2))
        sheet.write(ind, 6, final[i].mo_low.round(2))
        if final[i].percent_gain_loss > gainLossMargin:
            sheet.write(ind, 7, str(final[i].percent_gain_loss.round(2)) + '%', cell_format_green)
        elif final[i].percent_gain_loss < -gainLossMargin:
            sheet.write(ind, 7, str(final[i].percent_gain_loss.round(2)) + '%', cell_format_red)
        else:
            sheet.write(ind, 7, str(final[i].percent_gain_loss.round(2)) + '%')
    
    sheet.write(len(final) + 2, 0, 'Date')
    sheet.write(len(final) + 2, 1, dt.now().strftime('%m/%d/%Y'))
    sheet.autofit()
    xls.close()
    

def main():
    last_sat_date = dt.now() + relativedelta(weekday=SA(-1))  
    last_monday_date = dt.now() + relativedelta(weekday=MO(-1))
    last_monday_date = last_monday_date.strftime('%Y-%m-%d')
    start3mo = last_sat_date - relativedelta(months=3)
    last_sat_date = last_sat_date.strftime('%Y-%m-%d')
    start_3mo_date = start3mo.strftime('%Y-%m-%d')
    stock3mo = get_stock_data(list(tickersDict.keys()), start_3mo_date, last_sat_date)
    stock1wk = stock3mo.loc[last_monday_date:last_sat_date]
    stock_1wk_high, stock_1wk_low = get_highs_lows(stock1wk)
    stock_3mo_high, stock_3mo_low = get_highs_lows(stock3mo)
    friday_price = stock1wk.iloc[-1]['Close']
    percent_gain_loss = get_percent_gain_loss(list(tickersDict.values()), friday_price)
    create(tickersDict, friday_price, stock_1wk_high, stock_1wk_low, stock_3mo_high, stock_3mo_low, percent_gain_loss)
    export_to_excel()


if __name__ == '__main__':
    main()

