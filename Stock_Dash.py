import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import xlwings as xw
import yfinance
import statsmodels.formula.api as sm

sns.set_theme()

def dash_board(): 
    dash = xw.Book.caller().sheets[2]
    
    ticker = dash.range('C6').value
    bench = dash.range('K16').value
    start = dash['F6'].value
    end = dash['I6'].value
    
    data1 = yfinance.download(ticker, start = start, end=end)[['Adj Close']]
    data2 = yfinance.download(bench, start = start, end=end)[['Adj Close']]
    
    df = pd.merge(data1, data2, on = 'Date')
    
    df[str(ticker) + ' Normed_Return'] = df['Adj Close_x']/df.iloc[0]['Adj Close_x']
    df[str(bench) + ' Normed_Return'] = df['Adj Close_y']/df.iloc[0]['Adj Close_y']
    
    fig = plt.figure()
    df[str(ticker) + ' Normed_Return'].plot.line(ylabel = "Normed Return Scale", 
                  title = 'How Our Indvidual Asset Performs Against The Market Benchmark')
    df[str(bench) + ' Normed_Return'].plot.line()
    plt.legend([ticker + " Normed Return", bench + ' Normed Return'])
    dash.pictures.add(fig, update = True, name = 'Dual Chart', 
                  left=dash.range('E16').left, top=dash.range('E16').top);
    
    df.columns = [str(ticker) + '_Adj_Close', str(bench) + '_Adj_Close', str(ticker) + ' Normed_Return', 
                  str(bench) + ' Normed_Return']
    
    
    stock = str(ticker) + '_Adj_Close'
    market = str(bench) + '_Adj_Close'
    
    reg = sm.ols(formula = stock + '~' + market, data = df).fit()
    
    lower = reg.conf_int()[0][1]
    higher = reg.conf_int()[1][1]
    
    obs = reg.nobs
    alpha = reg.params[0]
    beta = reg.params[1]
    
    p_value = reg.pvalues[1]
    t_value = reg.tvalues[1]
    
    corr = np.corrcoef(df[stock], df[market])[1][0]
    
    dash['J37'].value = round(obs)
    dash['J38'].value = corr
    dash['J39'].value = beta
    dash['J40'].value = reg.rsquared
    dash['J41'].value = t_value
    dash['J42'].value = p_value
    dash['J43'].value = lower
    dash['J44'].value = higher
    dash['J45'].value = alpha
    
    
    fig = plt.figure()
    df[stock].plot.line(xlabel = 'Date', ylabel = ticker + " Adj Close Price", 
                   title = 'Individual Asset Performance')
    dash.pictures.add(fig, update = True, name='Chart', 
                 left=dash.range('E7').left, top=dash.range('E7').top);
    
    data1 = yfinance.download(ticker, start, end)
    first = data1.iloc[0]['Open']
    high = data1['Adj Close'].max()
    low = data1['Adj Close'].min()
    last = data1.iloc[-1]['Adj Close']
    pct_change = ((last - first) / first)*100
    
    dash['I9'].value = first
    dash['I10'].value = high
    dash['I11'].value = low
    dash['I12'].value = last
    dash['J13'].value = str(pct_change) + '%'
    
    
    fig = plt.figure(figsize=[5, 5])
    s = sns.regplot(x = market, y = stock, data = df, line_kws={"color": "darkgreen"})
    s.set(xlabel = 'Market Benchmark Price', ylabel = 'Individual Asset Price', 
          title = "Beta: Market Price as a Predictor for Individual Asset Price")
    dash.pictures.add(fig, update = True, name='Reg Chart', 
                      left = dash.range('E35').left, top = dash.range('E35').top);
    
    
    
dash_board()


def data_import(): 
    
    dash = xw.Book.caller().sheets[2]
    data = xw.Book.caller().sheets[1]
    ticker = dash.range('C6').value
    
    start = dash['F6'].value
    end = dash['I6'].value

    df = yfinance.download(ticker, start, end)
    
    data.range('A1').options(pd.DataFrame).value = df
    
data_import()
    
    
    
    
    
