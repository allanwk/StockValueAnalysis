StockValueAnalysis
==================

The scripts in this repository automate part of the process of estimating the value of stocks.  
This project has a scraper (stockInfo.py) that checks stock metrics on Yahoo Finance, using  
the Google Sheets API for storing the list of tracked stock symbols and the scripts output in
the cloud.

Os scripts neste repositório automatizam parte do processo de estimar o valor de ações.  
O projeto tem um scraper (stockInfo.py) que captura dados do Yahoo Finance, usando a API  
do Google Sheets para armazenar na nuvem a lista de ações rastreadas e a saída do script.

Technologies
------------

- BeautifulSoup
- Selenium
- Google Sheets API
- Pandas

Detailed Functionality
----------------------

The screener script (StockValueAnalysis.py) filters the stocks from theinput spreadsheet,  
using the `MASTER_SPREADSHEET_ID` environment variable, screening the tickers based on the information  
parsed by StockInfo.py. After this, in the output spreadsheet, `OUTPUT_SPREADSHEET_ID`, remain only the  
stocks classified as undervalued. Further analysis can be achieved using the ValueCalc.py script,  
that helps get more information about those companies and execute the intrinsic value calculation.