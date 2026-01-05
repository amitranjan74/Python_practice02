from backtesting import Backtest, Strategy

import pandas as pd
import time

import yfinance as yf

stocks=['AMZN','GOOG']

def fetch_data(stocks):
    data = yf.download(stocks, period='5d',interval='1m')
    return data



class ORBStrategy(Strategy):
    def init(self):
        self.orb_high = None
        self.orb_low = None
        self.trade = 0
        

    def next(self):
        # print(self.data.df)
        # time.sleep(1)

        # if len(self.data) < 2:
        #     return
        # else:
            self.current_time = self.data.index[-1].time()
            if self.current_time == pd.Timestamp("10:01").time():
                df = self.data.df
                d = pd.to_datetime(self.data.index[-1].date())
                # d = pd.to_datetime(d.date)
                df = df[df.index >= d]
                self.orb_high = df.High.max()
                self.orb_low = df.Low.min()
                self.trade = 0
                print(self.orb_high, self.orb_low)
               

            if not self.position and self.orb_high and self.orb_low:
                 
                 print('inside condition')
                 if (self.data.Close[-1] > self.orb_high) and self.trade == 0:
                      print('Buy condition satisfied')
                      self.buy()
                      self.trade = 1
                 elif (self.data.Close[-1] < self.orb_low) and self.trade == 0:
                      print('Sell condition satisfied')
                      self.sell()
                      self.trade = 1

            elif self.position:
                 print("I have some postion")
                 
                 if self.data.index[-1].time() == pd.Timestamp('15:20').time():
                      self.position.close()
                 
                 print(self.position)             
                 
                 



result = {}

for stock in stocks:
    data = fetch_data(stock)
    data.reset_index(inplace=True)
    data['Date']=data['Datetime'].dt.tz_localize(None)
    data.set_index('Date', inplace=True)
    print(data)
    bt = Backtest(data, ORBStrategy, cash=100_000, commission=.002)
    stats = bt.run()
    result[stock] = stats
    bt.plot()
