import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
# %matplotlib inline
import yfinance
import datetime as dt


def main(): 
    book = xw.Book.caller().sheets[0]
    
    symbol = book['C7'].value
    start = dt.date.isoformat(book['E7'].value)
    end = dt.date.isoformat(book['G7'].value)
    
    df = yfinance.download(symbol, start, end)[['Adj Close']]
    df['EWMA'] = df['Adj Close'].ewm(span = 26).mean()
    df['MACD'] = df['Adj Close'].ewm(span = 12).mean() - df['EWMA']
    df['Signal'] = df['MACD'].ewm(span = 20).mean()
    
    plt.style.use('seaborn')
    
    buy_sig = []
    sell_sig = []
    flag = -1
    
    for i in range(len(df)): 
        if df['MACD'][i] > df['Signal'][i]:
            if flag != 1:
                buy_sig.append(df['Adj Close'][i])
                sell_sig.append(np.nan)
                flag = 1
            else: 
                buy_sig.append(np.nan)
                sell_sig.append(np.nan)
        elif df['MACD'][i] < df['Signal'][i]:
            if flag != 0: 
                buy_sig.append(np.nan)
                sell_sig.append(df['Adj Close'][i])
                flag = 0
            else: 
                buy_sig.append(np.nan)
                sell_sig.append(np.nan)
        else: 
            buy_sig.append(np.nan)
            sell_sig.append(np.nan)
    
    
    df['Buy'] = buy_sig
    df['Sell'] = sell_sig
    
    fig, axes = plt.subplots(2, figsize=[12, 4])
    axes[0].plot(df.index, df['Signal'])
    axes[0].plot(df.index, df['MACD'])
    axes[0].legend(['Signal', 'MACD'])
    plt.title('MACD/Signal Technical Indicator')
    axes[1].plot(df.index, df['Adj Close'], c = 'purple', alpha = .3, label = 'Adj Close')
    axes[1].scatter(df.index, df['Buy'], marker = '^', c = 'green', label = 'Buy')
    axes[1].scatter(df.index, df['Sell'], marker = 'v', c = 'red', label = 'Sell')
    plt.tight_layout()
    plt.legend()
    plt.xlabel('Date')
    plt.ylabel('Stock Value')
    
    scale = book['O12']
    book.pictures.add(fig, name = 'Subplot', update = True, left = scale.left, top = scale.top)
    
    