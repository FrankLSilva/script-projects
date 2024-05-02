import config.constants as c
import pandas as pd
from gspread_dataframe import set_with_dataframe
from api import ConnectGoogleSheets
# Criando a conexão...
gs_connection = ConnectGoogleSheets()

# Exportando o DF da Viz 'North Stars'
## Acessando a worksheet desejada
worksheet = gs_connection.open_spreadsheet(c.SPREADSHEET).worksheet(c.WORKSHEET)
## Limpando os dados da worksheet
gs_connection.clear_worksheet(worksheet)
## Construindo o 'path' do arquivo a ser exportado
path_df = gs_connection.build_path('<NAME-FILE>','.csv')
## Salvando em DF para ser exportado
df = pd.read_csv(path_df)
## Salvando na spreadsheet
set_with_dataframe(worksheet, df)

