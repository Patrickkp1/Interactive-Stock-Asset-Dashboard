import pandas as pd
import numpy as np
from psaw import PushshiftAPI
import datetime as dt
import nltk
import xlwings as xw


def importing_reddit(): 
    dash = xw.Book.caller().sheets[2]
    
    start = (dt.datetime.now() - dt.timedelta(3)).timestamp()
    start = int(start)
    api = PushshiftAPI()
    
    sub = list(api.search_submissions(after = start, subreddit = 'Wallstreetbets', 
                                  filter = ['url', 'author', 'title', 'score', 'num_comments','selftext'], 
                                 limit =2000))
    outlist = []
    for s in sub:
        words = s.title.split()
        ticker = list(set(filter(lambda x: x.startswith('$'), words)))
        if len(ticker) > 0 and ticker[0][1:].isdigit() == False: 
            outlist.append(ticker[0][1:5])
            
    stocks = pd.DataFrame(outlist, columns = ['ticker'])
    
    import re
    l = []
    for k in outlist: 
        pattern = r'[0-9]'
        l.append(re.sub(pattern, '', k))
    
    outlist = list()
    for t in l: 
        if ',' in t:
            t = t.replace(',', '')
        if '.' in t: 
            t = t.replace('.', '')
        if '+' in t: 
            t = t.replace('+', '')
        if 'k' in t: 
            t = t.replace('k', '')
        if '$' in t:
            t = t.replace('$', '')
        if '?' in t:
            t = t.replace('?', '')
        
        regrex_pattern = re.compile(pattern = "["
                                    u"\U0001F600-\U0001F64F"
                                    u"\U0001F300-\U0001F5FF"  
                                    u"\U0001F680-\U0001F6FF"  
                                    u"\U0001F1E0-\U0001F1FF"  
                                    "]+", flags = re.UNICODE)
        t = regrex_pattern.sub(r'',t)
        outlist.append(t.upper())
    
    count = pd.DataFrame(outlist, columns = ['Ticker'])
    count = count[count['Ticker'] != '']
    c = count.value_counts().head(10)
    dash.range('O8').value = c

importing_reddit()

