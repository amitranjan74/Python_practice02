# %%
import yfinance as yf
import pandas as pd
import numpy as np
import mplfinance as mpf
from datetime import datetime, timedelta

# %%
# data = yf.download("QUESS.BO", start= '2022-03-31', end= '2024-09-02', interval= '15m')
data = yf.download("QUESS.BO", period = '1mo', interval= '15m')
data.shape

# %%
import pandas_ta as ta

# %%
data['max_5'] = data['Close'].rolling(window=5).max()
data

# %%
# 2. Calculate the 6 days ago Max(120, latest Close)
data['max_120_6_days_ago'] = data['Close'].shift(6).rolling(window=120).max()
data.tail()

# %%
data['num'] = [i for i in range(len(data))]

# %%
data['moving_avg_126'] = data['close'].rolling(window = 126).max()

# %%
data['cond1'] = np.where(data['max_5'] > (data['max_120_6_days_ago']),1,0)

# %%
data['sma_vol'] = ta.sma(data['volume'], length=5)

# %%
data['cond2'] = np.where(data['volume'] > (data['sma_vol']),1,0)

# %%
data['cond3'] = np.where(data['close'] > (data['close'].shift(1)),1,0)

# %%
data['Final_condi'] = np.where((data['cond1'] == 1) &
                               (data['cond2'] == 1) &
                               (data['cond3'] == 1), 1, 0)

# %%
data.to_csv('demo1')

# %%
buy_signal = data[data['Final_condi'] == 1]

# %%
buy_signal

# %%
data.to_csv("demo1")

# %%
data['Buy_price_signal'] = data['Final_condi']*data['close']

# %%
data.to_csv("demo1")

# %%
data['Buy_price_signal'].replace(0, np.nan, inplace = True)

# %%
data

# %%
buy_signal_markers = mpf.make_addplot(
    data['Buy_price_signal'],  # Plot signals on the closing prices
    type='scatter',        # Type of plot (scatter for points)
    markersize=100,        # Size of the markers
    marker='^',            # Marker type (up arrow for buy signals)
    color='g'              # Color of the marker (green)
)

# %%
mpf.plot(data,
         type='candle', 
         addplot = [buy_signal_markers],                      # Candlestick chart # Add the buy signal markers
         volume=True,                  # Include volume in the plot
         style='yahoo',                # Style of the plot (e.g., 'yahoo', 'classic', 'default')
         title='Candlestick Chart with Buy Signals',
         ylabel='Price',
         ylabel_lower='Volume')


