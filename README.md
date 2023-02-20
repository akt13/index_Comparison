## The code calculates your PnL if you have invested in NIFTYBEES instead of buying a different stock. 
## Also the top 5 stocks in up trend and down trend for the past 1 day, 3 days, 7 days, 2 weeks and a month from your Portfolio is displayed.

It works only for Zerodha
1. Upload the holdings export and Tradesbook exports to 'zerodha_exports' folder. Make sure both holding and tradesbook export are in the same time frame. Tradesbook export can be taken for 1yr interval each, no matter if the time interval overlaps; code handles it all.
2. Run the index_Comparison.py

To get the exports:
**Holdings Export**
Go to Zerodha Console > Portfolio > Holdings > Download: XLSX

**Tradebook Export**
Go to Zerodha Console > Reports > Tradebook > Specify the timeframe and export as excel

