from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep

def checkNa(var):
    try:
        var = float(var)
    except:
        var = float(0)
    return var

def getValue(selector, target_string, driver):
    elements = driver.find_elements_by_css_selector(selector)
    for index, element in enumerate(elements):
        if target_string in element.get_attribute("innerHTML"):
            return elements[index + 1].get_attribute("innerHTML")
            break

def findPB(driver):
    tds = driver.find_elements_by_css_selector('tr')
    pb_string = ""
    for i in range(len(tds)):
        try:
            tds = driver.find_elements_by_css_selector('tr')
            inner = tds[i].get_attribute("innerHTML")
            if "Price/Book" in inner:
                pb_string += inner
        except:
            tds = driver.find_elements_by_css_selector('tr')
            inner = tds[i].get_attribute("innerHTML")
            if "Price/Book" in inner:
                pb_string += inner
    pb_string = pb_string[pb_string.index('</td>'):]
    pb_string = pb_string[pb_string.index('">'):]
    pb_string = pb_string[2:pb_string.index('</td>')]
    return pb_string

def getStatistics(ticker, driver):
    ticker = ticker.upper()
    pe = 0
    pb = 0
    eps = 0

    yahoo_adress = "http://finance.yahoo.com/quote/{}.SA?p={}.SA".format(ticker, ticker)
    driver.get(yahoo_adress)
    driver.implicitly_wait(20)

    spans = driver.find_elements_by_css_selector('span')
    for i in range(len(spans)):
        try:
            spans = driver.find_elements_by_css_selector('span')
            inner = spans[i].get_attribute("innerHTML")
            if "Symbols similar to" in inner:
                return(0,0,0)
        except:
            spans = driver.find_elements_by_css_selector('span')
            inner = spans[i].get_attribute("innerHTML")
            if "Symbols similar to" in inner:
                return(0,0,0)

    pe = getValue("span", "PE Ratio (TTM)", driver)
    eps = getValue("span", "EPS (TTM)", driver)
    market_cap = getValue("span", "Market Cap", driver)

    yahoo_adress = "http://finance.yahoo.com/quote/{}.SA/key-statistics?p={}.SA".format(ticker, ticker)
    driver.get(yahoo_adress)
    driver.implicitly_wait(5)

    
    
    try:
        pb = findPB(driver)
    except:
        print("Informações estatísiticas indisponíveis para a ação {}".format(ticker))
        for i in range(180):
            if i % 10 == 0:
                print(i)
            sleep(1)
        try:
            driver.refresh()
            pb = findPB(driver)
        except:
            pb = 0
    

    return (pe,pb, eps, market_cap)