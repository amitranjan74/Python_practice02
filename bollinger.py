
from backtesting import Backtest, Strategy

import yfinance as yf
import pandas as pd
import time

import pandas_ta as ta

data = yf.download("RELIANCE.NS",start="2017-01-01",end="2023-04-30")

# data=data['2021-08-24':]


def SmaCross(closing_price, l):
    s = ta.sma(closing_price, l)
    return s

def bbands(close_price,l):
    b=ta.bbands(close_price,l)
    # print(b)
    return b[f'BBL_{l}_2.0'],b[f'BBU_{l}_2.0']


class bollingerbands(Strategy):
    l = 20
    stop_number = 0.05

    def init(self):

        self.sma20 = self.I(SmaCross, self.data.Close.s, 20)
        self.sma50 = self.I(SmaCross, self.data.Close.s, 50)
        self.sma200 = self.I(SmaCross, self.data.Close.s, 200)

        self.lowerbb, self.upperbb = self.I(bbands, self.data.Close.s, self.l)

    def next(self):
        if (self.data.Close[-1] < self.lowerbb[-1] and (self.data.Close[-2] >= self.lowerbb[-2])):
            
            p = self.data.Close[-1]
            
            if self.position.is_short:
                self.position.close()
            self.buy(sl=p-p*self.stop_number)
        elif (self.data.Close[-1] > self.upperbb[-1] and (self.data.Close[-2] <= self.upperbb[-2])
              and self.data.Close[-1] < self.sma20[-1]):

            p = self.data.Close[-1]

            if self.position.is_long:
                self.position.close()
            self.sell(sl=p+p*self.stop_number)



bt = Backtest(data, bollingerbands, 
              cash=100000, commission= 0.05)

output = bt.run()

print(output)

# bt.plot()

a=[0.01,0.02,0.03,0.04,0.05,0.06,0.07,0.08,0.09,0.1]
stats=bt.optimize(stop_number=a,maximize='Return [%]')
print(stats)
print(stats['_strategy'])
# bt.plot()
#overfitting    
