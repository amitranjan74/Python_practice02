
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta

def fetch_stock_data(ticker, period="6mo", interval="1d"):
    """
    Fetch stock data using yfinance
    """
    stock = yf.Ticker(ticker)
    df = stock.history(period=period, interval=interval)
    return df

def calculate_smma(series, length):
    """
    Calculate Smoothed Moving Average (SMMA)
    """
    smma = np.zeros_like(series)
    smma[0:length-1] = np.nan
    
    # Initialize with SMA
    smma[length-1] = np.mean(series[0:length])
    
    # Calculate SMMA for the rest
    alpha = 1 / length
    for i in range(length, len(series)):
        smma[i] = (series[i] - smma[i-1]) * alpha + smma[i-1]
    
    return smma

def calculate_z_smma(df, smma_length=12, z_length=30, source='Close'):
    """
    Calculate Z-Score of SMMA
    """
    # Calculate SMMA
    smma = calculate_smma(df[source].values, smma_length)
    
    # Calculate Z-Score of SMMA
    smma_series = pd.Series(smma)
    
    # Initialize arrays for mean and standard deviation
    mean = np.zeros_like(smma)
    std = np.zeros_like(smma)
    
    # Calculate rolling mean and std for Z-Score
    for i in range(z_length, len(smma)):
        window = smma[i-z_length+1:i+1]
        mean[i] = np.mean(window)
        std[i] = np.std(window)
    
    # Calculate Z-Score
    z_score = np.zeros_like(smma)
    z_score[:z_length] = np.nan
    mask = std != 0  # Avoid division by zero
    z_score[mask] = (smma[mask] - mean[mask]) / std[mask]
    
    return pd.Series(z_score, index=df.index)

def generate_signals(df, z_score, long_threshold=0.1, short_threshold=-0.1):
    """
    Generate trading signals based on Z-Score
    1 = Long, -1 = Short, 0 = Neutral
    """
    signals = pd.Series(0, index=df.index)
    
    # Set signals based on thresholds
    signals[z_score > long_threshold] = 1
    signals[z_score < short_threshold] = -1
    
    return signals

def scan_stocks(tickers, smma_length=12, z_length=30, long_threshold=0.1, short_threshold=-0.1):
    """
    Scan a list of stocks and categorize them based on Z-SMMA signals
    """
    results = {
        'buy': [],
        'sell': [],
        'hold': []
    }
    
    for ticker in tickers:
        try:
            # Fetch data
            df = fetch_stock_data(ticker)
            
            if len(df) < max(smma_length, z_length) + 10:
                print(f"Not enough data for {ticker}")
                continue
            
            # Calculate Z-SMMA
            z_score = calculate_z_smma(df, smma_length, z_length)
            
            # Generate signals
            signals = generate_signals(df, z_score, long_threshold, short_threshold)
            
            # Get current signal (most recent)
            current_signal = signals.iloc[-1]
            
            # Categorize
            if current_signal == 1:
                results['buy'].append(ticker)
            elif current_signal == -1:
                results['sell'].append(ticker)
            else:
                results['hold'].append(ticker)
            
            print(f"Processed {ticker}: Signal = {current_signal}")
            
        except Exception as e:
            print(f"Error processing {ticker}: {e}")
    
    return results

def calculate_extra_indicators(df):
    """
    Calculate additional indicators from the original PineScript
    """
    # ALMA calculation
    alma = df['Close'].rolling(window=8).mean()  # Simple approximation of ALMA
    
    # EMA calculations for plot lines
    ema3 = df['Close'].ewm(span=3, adjust=False).mean()
    ema19x2 = ema3.ewm(span=19*2, adjust=False).mean()
    ema19x3 = ema3.ewm(span=19*3, adjust=False).mean()
    ema15 = ema3.ewm(span=15, adjust=False).mean()
    
    return {
        'alma': alma,
        'ema3': ema3,
        'ema19x2': ema19x2,
        'ema19x3': ema19x3,
        'ema15': ema15
    }

def detailed_analysis(ticker, smma_length=12, z_length=30, long_threshold=0.1, short_threshold=-0.1):
    """
    Perform detailed analysis on a single stock
    """
    # Fetch data
    df = fetch_stock_data(ticker)
    
    # Calculate Z-SMMA
    z_score = calculate_z_smma(df, smma_length, z_length)
    
    # Generate signals
    signals = generate_signals(df, z_score, long_threshold, short_threshold)
    
    # Calculate extra indicators
    extra = calculate_extra_indicators(df)
    
    # Combine into results DataFrame
    results = pd.DataFrame({
        'Open': df['Open'],
        'High': df['High'],
        'Low': df['Low'],
        'Close': df['Close'],
        'Volume': df['Volume'],
        'SMMA_Z': z_score,
        'Signal': signals,
        'ALMA': extra['alma'],
        'EMA3': extra['ema3'],
        'EMA19x2': extra['ema19x2'],
        'EMA19x3': extra['ema19x3'],
        'EMA15': extra['ema15']
    })
    
    # Current signal status
    current_signal = signals.iloc[-1]
    signal_text = "BUY" if current_signal == 1 else "SELL" if current_signal == -1 else "HOLD"
    
    # Signal change detection
    signal_changes = []
    for i in range(2, len(signals)):
        if signals.iloc[i] == 1 and signals.iloc[i-1] == 1 and signals.iloc[i-2] == -1:
            signal_changes.append((signals.index[i], "LONG"))
        elif signals.iloc[i] == -1 and signals.iloc[i-1] == -1 and signals.iloc[i-2] == 1:
            signal_changes.append((signals.index[i], "SHORT"))
    
    return {
        'data': results,
        'current_signal': signal_text,
        'current_z_score': z_score.iloc[-1],
        'signal_changes': signal_changes
    }

if __name__ == "__main__":
    # Example stock list
    stock_list = [
    "ABCAPITAL.NS",  # Aditya Birla Capital
    "APOLLOHOSP.NS", # Apollo Hospitals
    "ASHOKLEY.NS",   # Ashok Leyland
    "AUROPHARMA.NS", # Aurobindo Pharma
    "BAJAJHLDNG.NS", # Bajaj Holdings
    "BALKRISIND.NS", # Balkrishna Industries
    "BANKINDIA.NS",  # Bank of India
    "BEL.NS",        # Bharat Electronics
    "CANBK.NS",      # Canara Bank
    "COFORGE.NS",    # Coforge
    "CONCOR.NS",     # Container Corporation of India
    "ESCORTS.NS",    # Escorts Kubota
    "FEDERALBNK.NS", # Federal Bank
    "GODREJPROP.NS", # Godrej Properties
    "HAL.NS",        # Hindustan Aeronautics
    "HINDPETRO.NS",  # Hindustan Petroleum
    "INDHOTEL.NS",   # Indian Hotels Company
    "INDUSTOWER.NS", # Indus Towers
    "MPHASIS.NS",    # Mphasis
    "PFC.NS"         # Power Finance Corporation
]
    
    # Set indicator parameters
    smma_length = 12
    z_length = 30
    long_threshold = 0.1
    short_threshold = -0.1
    
    print(f"Z-SMMA Stock Scanner")
    print(f"Parameters: SMMA Length={smma_length}, Z-Score Length={z_length}")
    print(f"Long Threshold={long_threshold}, Short Threshold={short_threshold}")
    print("-" * 50)
    
    # Scan stocks
    results = scan_stocks(stock_list, smma_length, z_length, long_threshold, short_threshold)
    
    # Print results
    print("\nScan Complete - Results Summary:")
    print("-" * 50)
    print(f"BUY Signals ({len(results['buy'])}):")
    for ticker in results['buy']:
        print(f"  - {ticker}")
    
    print(f"\nSELL Signals ({len(results['sell'])}):")
    for ticker in results['sell']:
        print(f"  - {ticker}")
    
    print(f"\nHOLD Signals ({len(results['hold'])}):")
    for ticker in results['hold']:
        print(f"  - {ticker}")
    
    # Detailed analysis for one stock example
    if len(results['buy']) > 0:
        sample_ticker = results['buy'][0]
    elif len(results['sell']) > 0:
        sample_ticker = results['sell'][0]
    elif len(results['hold']) > 0:
        sample_ticker = results['hold'][0]
    else:
        sample_ticker = stock_list[0]
    
    print("\nDetailed Analysis Example:")
    print("-" * 50)
    analysis = detailed_analysis(sample_ticker)
    print(f"Ticker: {sample_ticker}")
    print(f"Current Signal: {analysis['current_signal']}")
    print(f"Current Z-Score: {analysis['current_z_score']:.4f}")
    print(f"Recent Signal Changes:")
    for date, signal in analysis['signal_changes'][-3:]:
        print(f"  - {date.date()}: {signal}")
    
    # Example of how to save results
    print("\nSaving detailed data for", sample_ticker)
    analysis['data'].to_csv(f"{sample_ticker}_z_smma_analysis.csv")
    print(f"Data saved to {sample_ticker}_z_smma_analysis.csv")