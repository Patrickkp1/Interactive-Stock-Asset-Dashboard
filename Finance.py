import xlwings as xw
from bs4 import BeautifulSoup
import requests
import pandas as pd
import numpy as np
import yfinance
import datetime as dt


def importing_data(): 
    dash = xw.Book.caller().sheets[2]
    income = xw.Book.caller().sheets[3]
    
    ticker = dash['C6'].value
    start = dash['F6'].value
    end = dash['I6'].value
    
    
    myhead = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9", 
    "Accept-Encoding": "gzip, deflate, br", 
    "Accept-Language": "en-US,en;q=0.9", 
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36", 
    "X-Amzn-Trace-Id": "Root=1-60dbe4b1-259b415d2aff1e5c46208daf"
    }
    
    url = 'https://finance.yahoo.com/quote/' + ticker + '/financials?p=' + ticker
    page = requests.get(url, headers = myhead).content
    soup = BeautifulSoup(page, 'html.parser')
    s = soup.find_all('div', {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
    outlist = []
    for t in s:
        outlist.append(t.get_text(separator = '|').split('|'))
    inlist = []
    s = soup.find_all('div', {'class':'D(tbr) C($primaryColor)'})
    for t in s:
        inlist.append(t.get_text(separator = '|').split('|'))
        
        
        
    df1 = pd.DataFrame(outlist, columns = inlist[0]).set_index('Breakdown')
    cols = ["Total Revenue", "Cost of Revenue", "Gross Profit", "Operating Expense", 
       "Operating Income", "Pretax Income", "Diluted EPS",
       "Total Expenses", "Normalized Income", "EBIT", "Reconciled Cost of Revenue", 
        'Reconciled Depreciation', "Net Income from Continuing Operation Net Minority Interest",
        'Normalized EBITDA']
    
    if len(df1.columns) == 5:
        del df1[df1.columns[-1]]
    
    for col in df1.index:
        if col not in cols:
            df1.drop(col, inplace=True)
    
    income.range('A1').options(pd.DataFrame).value = df1
    
    
    balance = xw.Book.caller().sheets[4]
    
    url = 'https://finance.yahoo.com/quote/' + ticker + '/balance-sheet?p=' + ticker
    page = requests.get(url, headers = myhead).content
    soup = BeautifulSoup(page, 'html.parser')
    s = soup.find_all('div', {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
    outlist = []
    for t in s:
        outlist.append(t.get_text(separator = '|').split('|'))
    inlist = []
    s = soup.find_all('div', {'class':'D(tbr) C($primaryColor)'})
    for t in s:
        inlist.append(t.get_text(separator = '|').split('|'))
        
    
    df2 = pd.DataFrame(outlist, columns = inlist[0]).set_index('Breakdown')

    cols = ['Total Assets', 'Total Liabilities Net Minority Interest',
        'Total Equity Gross Minority Interest', 'Total Capitalization',
        'Common Stock Equity', 'Net Tangible Assets','Tangible Book Value',
            'Total Debt', 'Share Issued', 'Ordinary Shares Number']
    if len(df2.columns) == 4:
        del df2[df2.columns[-1]]
    
    for col in df2.index: 
        if col not in cols:
            df2.drop(col, inplace=True)
    
    balance.range('A1').options(pd.DataFrame).value = df2
    
    
    
    cash = xw.Book.caller().sheets[5]
    
    url = 'https://finance.yahoo.com/quote/' + ticker + '/cash-flow?p=' + ticker
    page = requests.get(url, headers = myhead).content
    soup = BeautifulSoup(page, 'html.parser')
    s = soup.find_all('div', {'class':'D(tbr) fi-row Bgc($hoverBgColor):h'})
    outlist = []
    for t in s:
        outlist.append(t.get_text(separator = '|').split('|'))
    inlist = []
    s = soup.find_all('div', {'class':'D(tbr) C($primaryColor)'})
    for t in s:
        inlist.append(t.get_text(separator = '|').split('|'))


    cols = ['Operating Cash Flow', "Investing Cash Flow", "Financing Cash Flow", 
       "End Cash Position", "Capital Expenditure",
        "Issuance of Capital Stock", "Free Cash Flow"]
    
    df3 = pd.DataFrame(outlist, columns = inlist[0]).set_index('Breakdown')

    if len(df3.columns) == 5:
        del df3[df3.columns[-1]]

    for col in df3.index: 
        if col not in cols:
            df3.drop(col, inplace=True)
            
    cash.range('A1').options(pd.DataFrame).value = df3
    
    
importing_data()
    
    
