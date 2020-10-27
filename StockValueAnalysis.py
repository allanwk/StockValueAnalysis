import pickle
import os.path
import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from stockInfo import *
from datetime import date, timedelta
from dotenv import load_dotenv

#Carregamento das variáveis de ambiente (IDs das planilhas do Google Sheets)
load_dotenv()

#Escopos de autorização da API do Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

driver = webdriver.Chrome(executable_path=r'C:\bin\chromedriver.exe')
start_at = ""

def main():
    #Autenticacao utilizando credenciais no arquivo JSON ou arquivo PICKLE
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

    #Inicializando o serviço sheets
    sheets_service = build('sheets', 'v4', credentials=creds)

    #Acessando valores da planilha tickers
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=os.environ.get("MASTER_SPREADSHEET_ID"),
        range='A2:A366',
        majorDimension='COLUMNS').execute()
    ticker_data = result.get('values', [])

    #Convertendo para Dataframe
    labels = ["Simbolo", "P/E", "P/B", "EPS", "P/E * P/B", "Preço", "Valor intrinseco", "Current BV", "Old BV", "Years", "1Y Dividends", "Market Cap Class"]
    ticker_df = pd.DataFrame(columns=labels)
    ticker_df["Simbolo"] = ticker_data[0]
    
    start_at = input("Começar de uma empresa específica? (Simbolo):")

    for index, row in ticker_df.iterrows():
        #Buscando empresa específica para começar
        if start_at == row["Simbolo"]:
            start_at = ""
        if start_at != "" and start_at != row["Simbolo"]:
            continue
        
        print("Buscando valores: P/E, P/B, EPS")
        print("Empresa {}/{}".format(index+1, len(ticker_df.index)))
        pe,pb,eps,market_cap = getStatistics(row["Simbolo"], driver)
        ticker_df.at[index, "P/E"] = checkNa(pe)
        ticker_df.at[index, "P/B"] = checkNa(pb)
        ticker_df.at[index, "EPS"] = checkNa(eps)
        ticker_df.at[index, "P/E * P/B"] = ticker_df.at[index, "P/E"] * ticker_df.at[index, "P/B"]
        
        #Classificação da empresa com relação à capitalização de mercado
        multiplier = market_cap[-1]
        market_cap = float(market_cap[:-1])
        if multiplier == "M":
            market_cap *= 1000000
        elif multiplier == "B":
            market_cap *= 1000000000

        market_cap_class = ""
        if market_cap <= 1700000000:
            market_cap_class = "Micro-cap"
        elif market_cap <= 11300000000:
            market_cap_class = "Small-cap"
        elif market_cap <= 56000000000:
            market_cap_class = "Mid-cap"
        else:
            market_cap_class = "Large-cap"

        ticker_df.at[index, "Market Cap Class"] = market_cap_class

        t = row["Simbolo"] + '.SA'
        yesterday = date.today() - timedelta(days=1)
        try:
            pr = data.DataReader(t, 'yahoo', yesterday - timedelta(days=1), yesterday)
            ticker_df.at[index, "Preço"] = pr.tail(1)["Close"]
        except:
            ticker_df.at[index, "Preço"] = 0
        if (float(row["P/B"]) == 0.0) or (float(row["P/E"]) == 0.0 and float(row["EPS"]) == 0.0):
                print("Refazendo buscas")
                print("Empresa {}/{}".format(index+1, len(ticker_df.index)))
                pe,pb,eps,market_cap = getStatistics(row["Simbolo"], driver)
                ticker_df.at[index, "P/E"] = checkNa(pe)
                ticker_df.at[index, "P/B"] = checkNa(pb)
                ticker_df.at[index, "P/E * P/B"] = ticker_df.at[index, "P/E"] * ticker_df.at[index, "P/B"]
                ticker_df.at[index, "EPS"] = checkNa(eps)
                ticker_df.at[index, "Market Cap Class"] = market_cap
                os.system('cls')
        ticker_df.to_excel('output_backup.xls')
    

    """Carregando dataframe do backup local, para caso a análise tenha sido fragmentada,
    para salvar os resultados na nuvem, já filtrando apenas ações sub-valorizadas"""

    ticker_df = pd.read_excel(r'C:\Users\allan\Documents\StockValueAnalysis\output_backup.xls', usecols="B:M")
    ticker_df = ticker_df.sort_values(by="P/E * P/B")
    ticker_df = ticker_df.fillna(0)
    values = []
    for index, row in ticker_df.iterrows():
        if row["EPS"] >= 0 and row["P/E * P/B"] <= 22.5:
            values.append(list(row))
    body = {"values": values}
    result = sheets_service.spreadsheets().values().update(
        spreadsheetId = os.environ.get("OUTPUT_SPREADSHEET_ID"),
        range='A2:L366',
        valueInputOption="RAW",
        body=body).execute()
    print('{0} celulas atualizadas na planilha na nuvem.'.format(result.get('updatedCells')))
    driver.close()

if __name__ == '__main__':
    main()