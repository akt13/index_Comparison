import os
import openpyxl
import pandas as pd
import yfinance as yf
from glob import glob
from datetime import timedelta, datetime


require_cols = ['Symbol', 'Trade Date','Trade Type','Quantity','Price','Order ID','Order Execution Time']

df = pd.concat([pd.read_excel(f, skiprows=14, usecols=require_cols) for f in glob("zerodha_exports/tradebook-*.xlsx")], ignore_index=True)
df.drop_duplicates('Order ID', inplace=True)
df = df[df['Trade Type'] != 'sell']
df.reset_index(inplace=True)

holdings_file = glob("zerodha_exports/holdings-*.xlsx")[0]
wb = openpyxl.load_workbook(holdings_file)
sh = wb.active
invested_val = sh['C15'].value
present_val = sh['C16'].value
unrealised_percentage = sh['C18'].value

def find_indexPrice(row):
    tradeDate = str(row['Trade Date'])
    ticker = row['Symbol'] + '.NS'
    start = datetime.strptime(tradeDate, '%Y-%m-%d')
    end = start + timedelta(days=1)
    index_Price = int(yf.download('NIFTYBEES.NS', start, end)['Adj Close'][0])
    return index_Price

df["index_ClosePrices"] = [find_indexPrice(row) for _, row in df.iterrows()]

LTP_index = yf.Ticker('NIFTYBEES.NS').history(period='1d')
index_LTP = LTP_index['Close'][0]

df["invest_Value"], df["index_P&L"] = zip(*[(row['Price']*row['Quantity'], index_LTP*(row['Price']*row['Quantity'])/row['index_ClosePrices'] - row['Price']*row['Quantity']) for _, row in df.iterrows()])

#df.to_excel("output.xlsx")

def pnl_Comparison():
    index_PnL = df['index_P&L'].sum()
    index_PnL_Percentage = (index_PnL/invested_val)*100
    print("Invest Value:", invested_val)
    print("Stock P&L:", present_val)
    print("Index P&L:", index_PnL)
    print("Stock P&L Percentage:", unrealised_percentage)
    print("Index P&L Percentage:", index_PnL_Percentage)

pnl_Comparison()
