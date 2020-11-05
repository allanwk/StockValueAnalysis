"""Este script é responsável por auxiliar na determinação do valor intrínseco
das ações encontradas na planilha de saída do script StockValueAnalysis.py.
Baseia-se nos gráficos salvos pelo usuário no diretório graphs na raíz do
projeto.
"""

import glob
import matplotlib.pyplot as plt
import pickle
import os.path
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys



#Carregamento das variáveis de ambiente (IDs das planilhas do Google Sheets)
load_dotenv()

#Escopos de autorização da API do Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

#Lista que irá armazenar coordenadas de cliques do mouse
coords = []

def focus(name, driver):
    """Função para auxiliar o foco em um elemento ao utilizar o Selenium"""
    driver.execute_script("document.getElementsByName('{}')[0].click()".format(name))

def onclick(event):
    """Registra a posição dos cliques nos gráficos apresentados"""
    if event.xdata != None and event.ydata != None:
        coords.append(event.ydata)
        if len(coords) == 4:
            plt.close()

def main():
    #Autenticação Google Sheets
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    sheets_service = build('sheets', 'v4', credentials=creds)

    #Acessando informações da planilha
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=os.environ.get("OUTPUT_SPREADSHEET_ID"),
        range='A2:N366',
        majorDimension='COLUMNS').execute()
    ticker_data = result.get('values', [])

    #Convertendo para DataFrame
    labels = ["Simbolo", "P/E", "P/B", "EPS", "P/E * P/B", "Preço", "Valor intrinseco", "Current BV", "Old BV", 
              "Years", "1Y Dividends", "Market Cap Class", "Price / Intrinsic Value", "Average Change in Book Value"]
    ticker_df = pd.DataFrame(columns=labels)
    for i in range(12):
        print(labels[i])
        ticker_df[labels[i]] = ticker_data[i]
    ticker_df["Price / Intrinsic Value"] = ticker_data[-2]
    ticker_df["Average Change in Book Value"] = ticker_data[-1]
    ticker_df = ticker_df.fillna(float(0))
    ticker_df = ticker_df.set_index("Simbolo")

    """Análise da evolução dos valores Book Value per Share e Debt to Equity Ratio de
    cada empresa no diretório graphs"""
    graphs = glob.glob("./graphs/*.png")
    for graph in graphs:
        ticker = graph[graph.index('\\') + 1:graph.index('.png')]
        coords.clear()
        print(ticker)
        img = plt.imread(graph)
        implot = plt.imshow(img)
        cid = implot.figure.canvas.mpl_connect('button_press_event', onclick)
        plt.show()

        val_scale = float(input("Entre com a escala: "))
        low1 = float(input("Valor inferior 1: "))
        low2 = float(input("Valor inferior 2: "))
        years = float(input("Entre com os anos decorridos: "))
        
        if len(coords) == 4:
            lower = coords[0]
            upper = lower - 50
            px_scale = lower-upper
            unit = val_scale / px_scale
            v1 = ((lower - coords[1])*unit)+low1
            lower = coords[2]
            upper = lower - 50
            px_scale = lower-upper
            unit = val_scale / px_scale
            v2 = ((lower - coords[3])*unit)+low2
            try:
                
                ticker_df.at[ticker, "Current BV"] = round(v2, 2)
                ticker_df.at[ticker, "Old BV"] = round(v1, 2)
                ticker_df.at[ticker, "Years"] = years
                print("Old BV: {}".format(v1))
                print("Current: {}".format(v2))
            except:
                print("Simbolo não encontrado, nome da imagem incorreto.")
        else:
            print("Não foram delimitados pontos suficientes")

    #Utilizar o site Buffet's Books para calcular o valor intrínseco
    driver = webdriver.Chrome(executable_path=r'C:\bin\chromedriver.exe')
    driver.get("https://www.buffettsbooks.com/how-to-invest-in-stocks/intermediate-course/lesson-21/")
    driver.implicitly_wait(10)
    fed_note = str(input("Entre com o valor do tesouro direto de 10 anos (68% = 0.68): "))
    ticker_df = ticker_df.fillna(0)
    ticker_df["Valor intrinseco"] = [float(0)]*len(ticker_df.index)

    for index, row in ticker_df.iterrows():
        currentBV = str(row["Current BV"])
        if currentBV == "0.0" or currentBV == "0":
            continue

        oldBV = str(row["Old BV"]).replace(',', '.')
        yrs = str(row["Years"]).replace(',', '.')
        dividends = str(row["1Y Dividends"]).replace(',', '.')

        names = ['cbv', 'obv', 'years', 'bvc', 'coupon', 'par', 'year', 'r']
        values_to_send = [currentBV, oldBV, yrs, "", dividends, currentBV, "10", fed_note]
        info = pd.DataFrame(data={'value': values_to_send}, index=names)

        driver.implicitly_wait(10)
        for name, info_row in info.iterrows():
            elm = driver.find_element_by_name(name)
            focus(name, driver)
            elm.send_keys(Keys.CONTROL + "a")
            elm.send_keys(Keys.DELETE)
            elm.send_keys(info_row["value"])
            if name == 'years':
                for inp in driver.find_elements_by_css_selector('input'):
                    if inp.get_attribute("value") == "Calculate":
                        inp.click()
                        break
                book_value_change = driver.find_element_by_name('totals').get_attribute('value')
                info.at['bvc', 'value'] = book_value_change.replace(',', '.')
                ticker_df.at[index, "Average Change in Book Value"] = round(float(book_value_change), 2)
        c = 0
        for inp in driver.find_elements_by_css_selector('input'):
            if inp.get_attribute("value") == "Calculate":
                if c == 1:
                    inp.click()
                    break
                else:
                    c+=1
                
        intrinsic = driver.find_element_by_name('total').get_attribute('value')
        ticker_df.at[index, "Valor intrinseco"] = round(float(intrinsic), 2)
        ticker_df.at[index, "Price / Intrinsic Value"] = round(float(str(row["Preço"]).replace(',', '.')) / float(intrinsic), 2)
    driver.close()

    print(ticker_df)
    """Preparando informações para o armazenamento na nuvem, filtrando 
    apenas as empresas que realmente foram analisadas"""
    ticker_df = ticker_df.fillna(float(0))
    values = [[index] + list(row) for index, row in ticker_df.iterrows() if int(row["Years"]) != 0]
    body = {"values": values}
    
    #Limpando a planilha na nuvem
    result = sheet.values().clear(
        spreadsheetId=os.environ.get("OUTPUT_SPREADSHEET_ID"),
        range='A2:N366').execute()
    #Salvando os novos valores
    result = sheets_service.spreadsheets().values().update(
        spreadsheetId = os.environ.get("OUTPUT_SPREADSHEET_ID"),
        range='A2:N366',
        valueInputOption="RAW",
        body=body).execute()
    print('{0} celulas atualizadas na planilha na nuvem.'.format(result.get('updatedCells')))
    
if __name__ == '__main__':
    main()