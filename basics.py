
from backtesting import Backtest, Strategy
from backtesting.lib import crossover

from backtesting.test import SMA, GOOG
import yfinance as yf

data = yf.download('BTC-USD',period='5y')


class SmaCross(Strategy):
    n1 = 10
    n2 = 20

    def init(self):
        close = self.data.Close
        self.smafast = self.I(SMA, close, self.n1)
        self.smaslow = self.I(SMA, close, self.n2)

    def next(self):
        if crossover(self.smafast, self.smaslow):
            self.buy()
        elif crossover(self.smaslow, self.smafast):
            self.sell()


bt = Backtest(data, SmaCross,
              cash= 100000, commission = 0.02,
              exclusive_orders = True)

output = bt.run()
print(output)
bt.plot()
print(output['_trades'].to_csv('trade.csv'))
        