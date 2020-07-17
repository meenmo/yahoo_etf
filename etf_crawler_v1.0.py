import os
import pandas as pd
from bs4 import BeautifulSoup
from urllib.request import urlopen


key = [
    'Previous Close',
    'Open',
    'Bid',
    'Ask',
    'Day\'s Range',
    '52 Week Range',
    'Volume',
    'Avg. Volume',
    'Net Assets',
    'NAV',
    'PE Ratio (TTM)',
    'Yield',
    'YTD Daily Total Return',
    'Beta (5Y Monthly)',
    'Expense Ratio (net)',
    'Inception Date'
    ]

while True:
    ticker_list = input("Ticker를 입력하세요: ").upper().split(',')
    ticker_dic  = {}

    try:
        for ticker in ticker_list:

            continue_dummy = 'Default'
            ticker = ticker.strip()

            if ticker == '':
                print(ticker + '  apple ')
                continue_dummy = 'y'
                continue
            
            url = 'https://finance.yahoo.com/quote/%s?p=%s&.tsrc=fin-srch' %(ticker, ticker)

            html = urlopen(url)
            bsObj  = BeautifulSoup(html, "html.parser")

            class_l = "W(100%)"
            text_l = bsObj.find("table",{"class":class_l})


            i = 0
            data_list_r = []
            for data in str(text_l).split('<td')[1:]:
                data = data.split('</span>')[0]

                if i == 1:
                    data = data.split('value">')[1].split('<')[0]

                else:
                    data = data.split('>')[-1]

                if data.split('>')[-1] in ['Day\'s Range', '52 Week Range']:
                    i = 1
                else:
                    i = 0

                data_list_r.append(data)
            
            key_temp = data_list_r[0::2]
            value = data_list_r[1::2]

            class_r = "D(ib) W(1/2) Bxz(bb) Pstart(12px) Va(t) ie-7_D(i) ie-7_Pos(a) smartphone_D(b) smartphone_W(100%) smartphone_Pstart(0px) smartphone_BdB smartphone_Bdc($seperatorColor)"
            text_r = bsObj.find("div",{"class":class_r})

            data_list_l = []
            for data in str(text_r).split('<span')[1:]:
                data = data.split('>')[1].replace('</span','')    
                data_list_l.append(data)


            key_temp += data_list_l[0::2]
            value += data_list_l[1::2]

            if key != key_temp:
                continue_dummy = 'y'
                continue

            ticker_dic[ticker] = value

        if (continue_dummy == 'Default')  or (ticker in ticker_list[-1]):
            table = pd.DataFrame(ticker_dic, index = key)
            wd = os.path.dirname(os.path.realpath(__file__)).replace('\\','/')
            filename = wd + '/' + "ETF_Yahoo.xlsx"
            table.to_excel(filename)
            os.system(filename)
            print('\n\n')
            continue

    except (TypeError, ValueError, EOFError):
        print("Ticker를 다시 입력해주세요", 'error')
        print('\n')
        pass

