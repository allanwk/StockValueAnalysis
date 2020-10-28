import bs4 as bs
import urllib.request
import requests
from time import sleep
import logging

logging.basicConfig(level=logging.INFO, filename="logfile.log")

def checkNa(var):
    try:
        var = float(var)
    except:
        var = float(0)
    return var

def getStatistics(ticker):
    ticker = ticker.upper()
    pe = 0
    pb = 0
    eps = 0
    price = 0
    market_cap = 0
    #source = urllib.request.urlopen("http://finance.yahoo.com/quote/{}.SA?p={}.SA".format(ticker, ticker)).read()
    source = requests.get("http://finance.yahoo.com/quote/{}.SA?p={}.SA".format(ticker, ticker)).text
    soup = bs.BeautifulSoup(source, 'lxml')
    spans = soup.find_all('span')
    for index, span in enumerate(spans):
        if "PE Ratio (TTM)" in span:
            pe = spans[index + 1].get_text()
        elif "EPS (TTM)" in span:
            eps = spans[index + 1].get_text()
        elif "Market Cap" in span:
            market_cap = spans[index + 1].get_text()
    #source = urllib.request.urlopen("http://finance.yahoo.com/quote/{}.SA/key-statistics?p={}.SA".format(ticker, ticker)).read()
    #Ao tentar obter o preço, caso esse não seja encontrado, a ação não está listada
    try:
        price = soup.find('span', attrs={"data-reactid":"32"}).get_text()
    except AttributeError:
        logging.warning("Ação {} não encontrada".format(ticker))
        return (pe,pb, eps, market_cap, price)

    source = requests.get("https://www.marketwatch.com/investing/stock/{}/company-profile?countrycode=br".format(ticker)).text
    soup = bs.BeautifulSoup(source, 'lxml')
    tds = soup.find_all('td')
    for index, td in enumerate(tds):
        if "Price to Book Ratio" in td:
            pb = tds[index + 1].get_text()
    if checkNa(pb) == 0:
        source = requests.get("https://finance.yahoo.com/quote/{}.SA/key-statistics?p={}.SA".format(ticker, ticker)).text
        soup = bs.BeautifulSoup(source, 'lxml')
        tds = soup.find_all('td')
        for index, td in enumerate(tds):
            if "Price/Book" in td.get_text():
                pb = tds[index + 1].get_text()

    return (pe,pb, eps, market_cap, price)