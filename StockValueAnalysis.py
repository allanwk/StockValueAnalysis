import pickle
import os.path
import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
import pandas_datareader.data as data
from stockInfo import *
from dotenv import load_dotenv
import progressbar

#Carregamento das variáveis de ambiente (IDs das planilhas do Google Sheets)
load_dotenv()

#Escopos de autorização da API do Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

#driver = webdriver.Chrome(executable_path=r'C:\bin\chromedriver.exe')
remove_list = []

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
    sheet = sheets_service.spreadsheets()

    #Acessando a planilha blacklist
    result = sheet.values().get(
        spreadsheetId=os.environ.get("BLACKLIST_SPREADSHEET_ID"),
        range='A1:A366',
        majorDimension='COLUMNS').execute()
    blacklist = {i: 0 for i in result.get('values', [])[0]}
    
    #Acessando valores da planilha tickers
    result = sheet.values().get(
        spreadsheetId=os.environ.get("MASTER_SPREADSHEET_ID"),
        range='A2:A366',
        majorDimension='COLUMNS').execute()
    ticker_data = result.get('values', [])

    #Convertendo para Dataframe
    labels = ["Simbolo", "P/E", "P/B", "EPS", "P/E * P/B", "Preço", "Valor intrinseco", "Current BV",
              "Old BV", "Years", "1Y Dividends", "Market Cap Class", "Price / Intrinsic Value", "Average Change in Book Value"]
    ticker_df = pd.DataFrame(columns=labels)
    ticker_df["Simbolo"] = ticker_data[0]

    bar = progressbar.ProgressBar(max_value=len(ticker_df.index))
    print("Obtendo informações sobre as ações.")

    for index, row in ticker_df.iterrows():
        bar.update(index)
        #Se a empresa estiver na planilha Blacklist, pular e remover do dataframe
        if row["Simbolo"] in blacklist:
            remove_list.append(index)
            continue

        #print("Empresa {}/{}".format(index+1, len(ticker_df.index)))
        pe, pb, eps, market_cap, price = getStatistics(row["Simbolo"])
        ticker_df.at[index, "P/E"] = checkNa(pe)
        ticker_df.at[index, "P/B"] = checkNa(pb)
        ticker_df.at[index, "EPS"] = checkNa(eps)

        #Se as métricas acima forem todas nulas, a ação não foi encontrada, portanto deve ser removida do Dataframe
        if (ticker_df.at[index, "P/E"] == 0 and ticker_df.at[index, "EPS"] == 0) or float(ticker_df.at[index, "P/E * P/B"]) == 0:
            remove_list.append(index)
            continue

        ticker_df.at[index, "P/E * P/B"] = ticker_df.at[index, "P/E"] * ticker_df.at[index, "P/B"]
        ticker_df.at[index, "Preço"] = checkNa(price)

        #Classificação da empresa com relação à capitalização de mercado
        try:
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
        except:
            print("Ação {} não encontrada.".format(row["Simbolo"]))

        ticker_df.to_excel('output_backup.xlsx')
    
    """Carregando dataframe do backup local, para caso a análise tenha sido fragmentada,
    para salvar os resultados na nuvem, já filtrando apenas ações sub-valorizadas"""

    #Limpando a planilha de saída no Google Sheets
    result = sheet.values().clear(
        spreadsheetId=os.environ.get("OUTPUT_SPREADSHEET_ID"),
        range='A2:N366').execute()

    ticker_df = pd.read_excel(r'C:\Users\allan\Documents\StockValueAnalysis\output_backup.xlsx', usecols="B:M")
    ticker_df = ticker_df.drop(remove_list)
    ticker_df = ticker_df.sort_values(by="P/E * P/B")
    ticker_df = ticker_df.fillna(0)

    values = [list(row) for index, row in ticker_df.iterrows() if row["EPS"] >= 0 and row["P/E * P/B"] <= 22.5]

    body = {"values": values}
    result = sheets_service.spreadsheets().values().update(
        spreadsheetId = os.environ.get("OUTPUT_SPREADSHEET_ID"),
        range='A2:N366',
        valueInputOption="RAW",
        body=body).execute()
    print('{0} celulas atualizadas na planilha na nuvem.'.format(result.get('updatedCells')))

if __name__ == '__main__':
    main()