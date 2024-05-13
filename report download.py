import requests
from openpyxl import Workbook
from openpyxl import load_workbook
from bs4 import BeautifulSoup
import pandas as pd
filesheet = "Envios de ws 2024 - Marzo.xlsx"
df=pd.read_excel(filesheet, sheet_name='SRX', usecols=['NFOLIO','ESTADO'])
df_folios = df.loc[(df['ESTADO'] == 'Digitado')|(df['ESTADO'] == 'En Ruta'), 'NFOLIO']
texto = "','".join(df_folios.astype(str))

texto = "'" + texto + "'"
#print(texto)

url = 'http://172.22.30.212/CONSULTA_F12_SRX_OMS_SAB_WMOS.php'
data = {
    #'select': 'f12',
    'PARAMETRO': texto
}
response = requests.post(url, data=data)

if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')
    tabla = soup.find('table')
    if tabla:
        #'''
        # Crear un nuevo libro de Excel
        libro_excel = Workbook()
        hoja_excel = libro_excel.active
        
        # Iterar sobre las filas y celdas de la tabla HTML y escribir en el libro de Excel
        for fila_html in tabla.find_all('tr'):
            fila_excel = []
            for celda_html in fila_html.find_all(['th', 'td']):
                fila_excel.append(celda_html.text.strip())
            print(fila_excel)
            hoja_excel.append(fila_excel)
        # Guardar el libro de Excel
        libro_excel.save('informacion.xlsx')
        #'''
        #print(tabla)
else:
    print('Error')
