import datetime as dt
import yfinance 
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
import numpy as np


def RSI_fn():
    book = xw.Book.caller().sheets[0]
#     path = '/Users/patrickpoleshuk/Desktop/Bollinger Bands/Bollinger_Bands.xlsm'
#     book = xw.Book(path).sheets[0]
    symbol = book['C7'].value
    start = book['E7'].value
    end = book['G7'].value
    
    df = yfinance.download(symbol, start, end)[['Adj Close']]

    df['delta'] = df['Adj Close'] - df['Adj Close'].shift(1)
    df['gain'] = np.where(df['delta'] > 0, df['delta'], 0)
    df['loss'] = np.where(df['delta'] < 0, abs(df['delta']), 0)
    
    avg_gain = []
    avg_loss = []
    
    gain = list(df['gain'])
    loss = list(df['loss'])
    
    n = 14
    for i in range(len(df)):
        if i < n:
            avg_gain.append(np.nan)
            avg_loss.append(np.nan)
        if i == n: 
            avg_gain.append(df['gain'].rolling(window = n).mean().to_list()[n])
            avg_loss.append(df['loss'].rolling(window = n).mean().to_list()[n])
        if i > n:
            avg_gain.append(((n-1)*avg_gain[i-1] + gain[i])/n)
            avg_loss.append(((n-1)*avg_loss[i-1] + loss[i])/n)
    
    buy = []
    sell = []
    
    df['avg_gain'] = np.array(avg_gain)
    df['avg_loss'] = np.array(avg_loss)
    df['RS'] = df['avg_gain'] / df['avg_loss']
    df['RSI'] = 100 - (100/(1 + df['RS']))
    
    for i in range(len(df)): 
        if df['RSI'][i] >= 70: 
            buy.append(np.nan)
            sell.append(df['Adj Close'][i])
        elif df['RSI'][i] <= 30: 
            buy.append(df['Adj Close'][i])
            sell.append(np.nan)
        else: 
            buy.append(np.nan)
            sell.append(np.nan)
    
    plt.style.use('seaborn')
    df['Buy'] = buy
    df['Sell'] = sell
    fig = plt.figure(figsize = [14, 8])
    df['Adj Close'].plot(alpha = .3, lw = 3)
    plt.scatter(df.index, df['Buy'], marker = '^', c = 'green', label = 'Buy')
    plt.scatter(df.index, df['Sell'], marker = 'v', c = 'red', label = 'Sell')
    
    scale = book['B41']
    book.pictures.add(fig, update = True, name = 'RSI', top = scale.top, left = scale.left)
    
    

