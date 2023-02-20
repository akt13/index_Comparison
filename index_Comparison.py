import os
import openpyxl
import pandas as pd
import yfinance as yf
from glob import glob
from datetime import timedelta, datetime, date


require_cols = ['Symbol', 'Trade Date', 'Trade Type',
                'Quantity', 'Price', 'Order ID', 'Order Execution Time']

df = pd.concat([pd.read_excel(f, skiprows=14, usecols=require_cols)
               for f in glob("zerodha_exports/tradebook-*.xlsx")], ignore_index=True)
df.drop_duplicates('Order ID', inplace=True)
df = df[df['Trade Type'] != 'sell']
df.reset_index(inplace=True)

holdings_file = glob("zerodha_exports/holdings-*.xlsx")[0]
wb = openpyxl.load_workbook(holdings_file)
sh = wb.active
invested_val = sh['C15'].value
present_val = sh['C16'].value
unrealised_percentage = sh['C18'].value
df_holding = pd.read_excel(holdings_file[0], skiprows=22, usecols=[
                           'Symbol', 'Average Price'])


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

df["invest_Value"], df["index_P&L"] = zip(*[(row['Price']*row['Quantity'], index_LTP*(
    row['Price']*row['Quantity'])/row['index_ClosePrices'] - row['Price']*row['Quantity']) for _, row in df.iterrows()])


def pnl_Comparison():
    index_PnL = df['index_P&L'].sum()
    index_PnL_Percentage = (index_PnL/invested_val)*100
    print("Invest Value:", invested_val)
    print("Stock P&L:", present_val)
    print("Index P&L:", index_PnL)
    print("Stock P&L Percentage:", unrealised_percentage)
    print("Index P&L Percentage:", index_PnL_Percentage)


pnl_Comparison()


def portfolio_Change(row):
    # Get historical data for the past 135 days
    end_date = date.today()
    start_date = end_date - timedelta(days=200)
    ticker = row['Symbol'] + '.NS'
    data = yf.download(ticker, start=start_date, end=end_date)

    # Calculate the price change percentage for the last x days
    one_days_data = data.tail(2)
    three_days_data = data.tail(4)
    seven_days_data = data.tail(8)
    fourteen_days_data = data.tail(15)
    thirty_days_data = data.tail(31)

    one_days_close = one_days_data['Close']
    one_days_change = (
        (one_days_close.iloc[1] - row['Average Price']) / row['Average Price']) * 100

    three_days_close = three_days_data['Close']
    three_days_change = (
        (three_days_close.iloc[1] - row['Average Price']) / row['Average Price']) * 100

    seven_days_close = seven_days_data['Close']
    seven_days_change = (
        (seven_days_close.iloc[1] - row['Average Price']) / row['Average Price']) * 100

    fourteen_days_close = fourteen_days_data['Close']
    fourteen_days_change = (
        (fourteen_days_close.iloc[1] - row['Average Price']) / row['Average Price']) * 100

    thirty_days_close = thirty_days_data['Close']
    thirty_days_change = (
        (thirty_days_close.iloc[1] - row['Average Price']) / row['Average Price']) * 100

    return one_days_change, three_days_change, seven_days_change, fourteen_days_change, thirty_days_change


oneDay_lst, threeDay_lst, sevenDay_lst, fourteenDay_lst, thirtyDay_lst = [], [], [], [], []
for x, row in df_holding.iterrows():
    one_days_change, three_days_change, seven_days_change, fourteen_days_change, thirty_days_change = portfolio_Change(
        row)
    oneDay_lst.append(one_days_change)
    threeDay_lst.append(three_days_change)
    sevenDay_lst.append(seven_days_change)
    fourteenDay_lst.append(fourteen_days_change)
    thirtyDay_lst.append(thirty_days_change)
df_holding["OneDay_Change"] = oneDay_lst
df_holding["ThreeDay_Change"] = threeDay_lst
df_holding["SevenDay_Change"] = sevenDay_lst
df_holding["FourteenDay_Change"] = fourteenDay_lst
df_holding["ThirtyDay_Change"] = thirtyDay_lst

df_holding.to_excel("output.xlsx")

df_holding = df_holding.sort_values(by=['OneDay_Change'], ascending=False)
print('Top 5 stocks in up-trend from last 1 day')
print(df_holding[['Symbol', 'OneDay_Change']].head(5))
print('-------------------------------------------')
print('Top 5 stocks in down-trend from last 1 day')
print(df_holding[['Symbol', 'OneDay_Change']].tail(5))
print('-------------------------------------------')
print('-----------------*********-----------------')

df_holding = df_holding.sort_values(by=['ThreeDay_Change'], ascending=False)
print('Top 5 stocks in up-trend from last 3 day')
print(df_holding[['Symbol', 'ThreeDay_Change']].head(5))
print('-------------------------------------------')
print('Top 5 stocks in down-trend from last 3 day')
print(df_holding[['Symbol', 'ThreeDay_Change']].tail(5))
print('-------------------------------------------')
print('-----------------*********-----------------')

df_holding = df_holding.sort_values(by=['SevenDay_Change'], ascending=False)
print('Top 5 stocks in up-trend from last 7 day')
print(df_holding[['Symbol', 'SevenDay_Change']].head(5))
print('-------------------------------------------')
print('Top 5 stocks in down-trend from last 7 day')
print(df_holding[['Symbol', 'SevenDay_Change']].tail(5))
print('-------------------------------------------')
print('-----------------*********-----------------')

df_holding = df_holding.sort_values(by=['FourteenDay_Change'], ascending=False)
print('Top 5 stocks in up-trend from last 2 weeks')
print(df_holding[['Symbol', 'FourteenDay_Change']].head(5))
print('-------------------------------------------')
print('Top 5 stocks in down-trend from last 2 weeks')
print(df_holding[['Symbol', 'FourteenDay_Change']].tail(5))
print('-------------------------------------------')
print('-----------------*********-----------------')

df_holding = df_holding.sort_values(by=['ThirtyDay_Change'], ascending=False)
print('Top 5 stocks in up-trend from last 1 month')
print(df_holding[['Symbol', 'ThirtyDay_Change']].head(5))
print('-------------------------------------------')
print('Top 5 stocks in down-trend from last 1 month')
print(df_holding[['Symbol', 'ThirtyDay_Change']].tail(5))
print('-------------------------------------------')
print('-----------------*********-----------------')
