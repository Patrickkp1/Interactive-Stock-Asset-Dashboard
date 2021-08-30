import xlwings as xw
from bs4 import BeautifulSoup
import requests
import pandas as pd
import numpy as np


def stats_import(): 
    dash = xw.Book.caller().sheets[2]
    income = xw.Book.caller().sheets[3]
    balancee = xw.Book.caller().sheets[4]
    cashh = xw.Book.caller().sheets[5]
    dataa = xw.Book.caller().sheets[1]
    ticker = dash['C6'].value
    myhead = {  
              "Accept-Encoding": "gzip, deflate, br", 
              "Accept-Language": "en-US,en;q=0.9", 
              "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML,like Gecko) Chrome/91.0.4472.114 Safari/537.36", 
              "X-Amzn-Trace-Id": "Root=1-60dbe4b1-259b415d2aff1e5c46208daf"
             }
    url = 'https://finance.yahoo.com/quote/' + ticker + '/key-statistics?p=' + ticker
    page = requests.get(url, headers = myhead).content
    soup = BeautifulSoup(page, 'html.parser')
    s = soup.find_all('div', {'class':'W(100%) Whs(nw) Ovx(a)'})
    outlist = list()
    for t in s:
        outlist.append(t.get_text(separator = '|').split('|'))
    inlist = outlist[0]
    for val in inlist: 
        if val == ' ': 
            inlist.remove(val)
    
    del inlist[0]
    del inlist[8]
    del inlist[15]
    del inlist[-1]
    del inlist[37]
    del inlist[29]
    del inlist[43]
    del inlist[50]
    del inlist[57]
    del inlist[64]
    del inlist[0]
    
    real = pd.DataFrame(np.reshape(inlist, (10, 7)))
    real[real.columns[6]] = real[real.columns[6]].shift(1)
    real.columns = real.iloc[0]
    real.set_index(real.columns[6], inplace = True)
    real.drop(real.index[0], inplace=True)
    income['A18'].options(pd.DataFrame).value = real
    s = soup.find_all('table', {'class':'W(100%) Bdcl(c)'})
    
    outlist = list()
    for t in s:
        outlist.append(t.get_text(separator = '|').split('|'))
    profit = outlist[4]
    returns = outlist[5]
    for val in returns: 
        if val == ' ': 
            returns.remove(val)
        
    returns_info = pd.DataFrame(np.reshape(returns, (2, 3)))
    profit.insert(2, '(ttm)')
    for val in profit: 
        if val == ' ': 
            profit.remove(val)
            
    profit = np.reshape(profit, (2, 3))
    profit_info = pd.DataFrame(profit)

    for val in outlist[6]: 
        if val == ' ': 
            outlist[6].remove(val)
    outlist[6].insert(13, '(ttm)')
    income_info = pd.DataFrame(np.reshape(outlist[6], (8, 3)))

    balance = outlist[7]
    for val in balance: 
        if val == ' ': 
            balance.remove(val)
    balance_info = pd.DataFrame(np.reshape(balance, (6, 3)))

    cash = outlist[8]
    for val in cash: 
        if val == ' ': 
            cash.remove(val)
    cash_info = pd.DataFrame(np.reshape(cash, (2, 3)))

    dividend = outlist[2]
    for val in dividend: 
        if val == ' ': 
            dividend.remove(val)
    dividend_info = pd.DataFrame(np.reshape(dividend, (10, 3)))

    price = outlist[0]
    try: 
        conds = [' ', '3']
        for val in price: 
            if val in conds:
                price.remove(val)
        price_info = pd.DataFrame(np.reshape(price, (7, 2)))

    except Exception: 
        conds = [' ', '3']
        for val in price: 
            if val in conds:
                price.remove(val)
        price_info = pd.DataFrame(np.reshape(price, (7, 2)))
    
    
    stats = outlist[1]
    try: 
        conds = [' ', '1', '3', '4', '5', '6']
        for val in stats: 
            if val in conds: 
                stats.remove(val)
        stats_info = pd.DataFrame(np.reshape(stats, (12, 2)))
    
    except Exception: 
        conds = [' ', '1', '3', '4', '5', '6']
        for val in stats: 
            if val in conds: 
                stats.remove(val)
        stats_info = pd.DataFrame(np.reshape(stats, (12, 2)))
    
    info = outlist[3]
    del info[4]
    info_df = pd.DataFrame(np.reshape(info, (2, 3)))

    dash['L1'].options(pd.DataFrame).value = info_df
    income['A29'].options(pd.DataFrame).value = profit_info
    dataa['I31'].options(pd.DataFrame).value = stats_info
    dataa['I3'].options(pd.DataFrame).value = price_info
    balancee['A14'].options(pd.DataFrame).value = balance_info
    cashh['A11'].options(pd.DataFrame).value = cash_info
    dataa['J13'].options(pd.DataFrame).value = dividend_info
    dataa['I26'].options(pd.DataFrame).value = returns_info
    income['F30'].options(pd.DataFrame).value = income_info
    
stats_import()
