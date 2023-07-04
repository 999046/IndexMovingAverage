from pandas_datareader import data as pdr
import pandas as pd
import yfinance as yf
import xlsxwriter
from datetime import date
from dateutil.relativedelta import relativedelta

today = date.today()
# yf.pdr_override() # <== that's all it takes :-)

# # download dataframe
# data = pdr.get_data_yahoo("AAPL", start = '2023-06-01', end = '2023-06-30')
# #data['sma_200'] = data['Close'].rolling(window=200).mean()
# data['sma_7'] = data['Close'].rolling(window=7).mean()
# #final_list = data.tolist()
# print(data['sma_7'].iloc[-1])
# #print(data[6:])

def getAllStockIndex():
   indexes = pd.read_html('https://ournifty.com/stock-list-in-nse-fo-futures-and-options.html#:~:text=NSE%20F%26O%20Stock%20List%3A%20%20%20%20SL,%20%201000%20%2052%20more%20rows%20')[0]
   indexes = indexes.SYMBOL.to_list()
   indexes.remove('FINNIFTY')
   indexes.remove('MIDCPNIFTY')
   indexes.remove('BANKNIFTY')
   indexes.remove('NIFTY')
   
   for count in range(len(indexes)):
     indexes[count] = indexes[count] + ".NS"
   return indexes

def get_moving_average(yf_data, indexes, no_of_days):
  m_avg_frame = pd.DataFrame({'Index': indexes})
  m_avg = []
  for count in range(len(tickers)):
    m_avg.append(yf_data['Close'][tickers[count]].rolling(window=no_of_days).mean().iloc[-1])
    #   print(tickers[count] + "," + str(yf_data['Close'][tickers[count]].rolling(window=7).mean().iloc[-1]))
  m_avg_frame['avg'] = m_avg 
  return m_avg_frame
 
tickers = getAllStockIndex() 

yf_data = yf.download(tickers, start = today- relativedelta(years=1), end = today)
m_avg_7day = get_moving_average(yf_data, tickers, 7)
m_avg_30day = get_moving_average(yf_data, tickers, 30)
m_avg_50day = get_moving_average(yf_data, tickers, 50)
m_avg_200day = get_moving_average(yf_data, tickers, 200)
writer = pd.ExcelWriter('movingAgerage.xlsx', engine='xlsxwriter')
m_avg_7day.to_excel(writer, sheet_name='7DayMovingAvg')
m_avg_30day.to_excel(writer, sheet_name='30DayMovingAvg')
m_avg_50day.to_excel(writer, sheet_name='50DayMovingAvg')
m_avg_200day.to_excel(writer, sheet_name='200DayMovingAvg')

writer.close()

