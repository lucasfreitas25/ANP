from urllib.request import urlretrieve 
from zipfile import ZipFile
import pandas as pd 
import openpyxl
from openpyxl.styles import Border, Side, Font
from openpyxl.utils import get_column_letter

url = 'https://www.gov.br/anp/pt-br/assuntos/producao-e-fornecimento-de-biocombustiveis/biometano/biometano-dados-abertos.zip'

arquivo = 'biometano-dados-abertos.zip'
urlretrieve(url, arquivo)
with ZipFile(arquivo, "r") as zip:
    zip.extractall('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIOMETANO\\Dados em csv')
    
    
df_capacidade = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIOMETANO\\Dados em csv\\Biometano_DadosAbertos_CSV_Capacidade.csv')
# df_capacidade['CNPJ'] = df_capacidade['CNPJ'].astype(str)
df_capacidade.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIOMETANO\\Dados Biometano Capacidade.xlsx', index=False)

df_prod = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIOMETANO\\Dados em csv\\Biometano_DadosAbertos_CSV_Producao.csv')
# df_prod['Produção de Biodiesel'] = df_prod['Produção de Biodiesel'].str.replace(',', '.').astype(float)
df_prod.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIOMETANO\\Dados Biometano Produção.xlsx', index=False)

#NÃO TEM NENHUM DADO DE MATO GROSSO POR ISSO NÃO FOI CONTINUADO
