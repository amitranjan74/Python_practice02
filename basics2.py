from backtesting import Backtest, Strategy
import pandas_ta as ta
import time


import yfinance as yf

data = yf.download('^NSEBANK', period='10y', interval='1d')

def SmaCross(closing_price, l):
    sma1 = ta.sma(closing_price, l)
    return sma1

def get_atr(high, low, close, l):
    atr1= ta.atr(high, low, close, l)
    return atr1

class SmaStrategy(Strategy):
    n1 = 15
    n2 = 20

    atr_length = 14

    def init(self):
        self.smafast = self.I(SmaCross, self.data.Close.s, self.n1)
        self.smaslow = self.I(SmaCross, self.data.Close.s, self.n2)
        self.atr = self.I(get_atr, self.data.High.s, self.data.Low.s, self.data.Close.s, self.atr_length)

    def next(self):
        if (self.smafast[-1] > self.smaslow[-1]) & (self.smafast[-2] <= self.smaslow[-2]):
            if self.position.is_short:
                self.position.close()
            self.buy(sl=self.data.Close[-1] - 2*self.atr[-1])
        elif (self.smafast[-1] < self.smaslow[-1]) & (self.smafast[-2] >= self.smaslow[-2]):
            if self.position.is_long:
                self.position.close()
            # self.sell(sl=self.data.Close[-1] + 2*self.atr[-1])


bt = Backtest(data, SmaStrategy,
               cash=100000, commission=0.02)

output = bt.run()
print(output)
bt.plot()

