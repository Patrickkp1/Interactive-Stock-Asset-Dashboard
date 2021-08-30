import yfinance 
import datetime as dt
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw

plt.style.use('seaborn')

def new_columns(): 
    book = xw.Book.caller().sheets[0]

#     path = '/Users/patrickpoleshuk/Desktop/Bollinger Bands/Bollinger_Bands.xlsx'
#     book = xw.Book(path).sheets[0]
    
    ticker = book.range('C7').value 
    start = book.range('E7').value
    end = book.range('G7').value
    symbol = 'Open'
    green = yfinance.download(str(ticker))[str(start):str(end)][[symbol]]
    green['EWMA'] = green[symbol].ewm(span = 20).mean()
    green['Lower'] = green['EWMA'] - 2*green[symbol].ewm(span = 20).std()
    green['Upper'] = green['EWMA'] + 2*green[symbol].ewm(span = 20).std()
    
    buy_sig = []
    sell_sig = []
        
    for i in range(len(green)):
        
        if green['Lower'][i] > green[symbol][i]: 
            buy_sig.append(green[symbol][i])
            sell_sig.append(np.nan)
            
        elif green['Upper'][i] < green[symbol][i]: 
            buy_sig.append(np.nan)
            sell_sig.append(green[symbol][i])
            
        else: 
            buy_sig.append(np.nan)
            sell_sig.append(np.nan)
    
    green['Buy'] = buy_sig
    green['Sell'] = sell_sig
    
    fig = plt.figure(figsize=[12, 10])
    green[symbol].plot(kind = 'line', alpha = .3)
    green['EWMA'].plot(kind = 'line', alpha = .3)
    green['Lower'].plot(kind = 'line', alpha = .3)
    green['Upper'].plot(kind = 'line', alpha = .3)
    
    plt.scatter(green.index, green['Buy'], label = 'Buy', marker='^', c = 'green')
    plt.scatter(green.index, green['Sell'], label = 'Sell', marker='v', c = 'red')
    
    plt.ylabel('Open Price')
    plt.title(str(symbol) + ' price of: ' + str(ticker))
    plt.legend()
    book.pictures.add(fig, name = 'Chart', update = True)

new_columns()
