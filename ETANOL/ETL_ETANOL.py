from urllib.request import urlretrieve 
from zipfile import ZipFile
import pandas as pd 
import openpyxl
from openpyxl.styles import Border, Side, Font
from openpyxl.utils import get_column_letter

#BAIXA O ARQUIVO E EXTRAI ELE PARA UMA PASTA
url = 'https://www.gov.br/anp/pt-br/assuntos/producao-e-fornecimento-de-biocombustiveis/etanol/arquivos-etanol/pb-da-etanol.zip'
arquivo = 'pb-da-etanol.zip'

urlretrieve(url, arquivo)

with ZipFile(arquivo, "r") as zip:
    zip.extractall('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados em csv')

#FAZ LIMPEZA E MUDANÇAS ESTRUTURAIS
df_capacidade = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados em csv\\Etanol_DadosAbertos_CSV_Capacidade.csv')
df_capacidade['CNPJ'] = df_capacidade['CNPJ'].astype(str)
df_capacidade.fillna(value=0, inplace=True)
df_capacidade.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados Etanol Capacidade.xlsx', index=False)

df_matprima = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados em csv\\Etanol_DadosAbertos_CSV_MatériaPrima.csv')
df_matprima['Quantidade Processada (t)'] = df_matprima['Quantidade Processada (t)'].str.replace(',', '.').astype(float)
df_matprima.fillna(value=0, inplace=True)
df_matprima.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados Etanol Materia Prima.xlsx', index=False)

df_prod = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados em csv\\Etanol_DadosAbertos_CSV_Produç╞o.csv')
df_prod.fillna(value=0, inplace=True)
df_prod.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados Etanol Produção.xlsx', index=False)

#JUNTAR TODAS AS PLANILHAS EM UMA
planilha_principal = openpyxl.Workbook()

wb_capacidade = openpyxl.load_workbook('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados Etanol Capacidade.xlsx')  
wb_matprima = openpyxl.load_workbook("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados Etanol Materia Prima.xlsx")  
wb_prod = openpyxl.load_workbook('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\Dados Etanol Produção.xlsx')


aba_capacidade = planilha_principal.create_sheet("CAPACIDADE")
aba_matprima = planilha_principal.create_sheet("MATERIA PRIMA")
aba_prod = planilha_principal.create_sheet("PRODUÇÃO")

for linha in wb_capacidade.active.iter_rows(values_only=True):
    aba_capacidade.append(linha)

for linha in wb_matprima.active.iter_rows(values_only=True):
    aba_matprima.append(linha)
    
for linha in wb_prod.active.iter_rows(values_only=True):
    aba_prod.append(linha)

for aba in planilha_principal.sheetnames:
    if aba not in ["CAPACIDADE", "MATERIA PRIMA", "PRODUÇÃO"]:
        del planilha_principal[aba]

########################
def ajustar_colunas(aba):
    for coluna in aba.columns:
        max_length = 0
        coluna = [cell for cell in coluna]
        for cell in coluna:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        aba.column_dimensions[get_column_letter(coluna[0].column)].width = adjusted_width

lista_abas = [aba_capacidade, aba_matprima, aba_prod]
for abas in lista_abas:
    ajustar_colunas(abas)
    
    
############################    
worksheet = planilha_principal.active
df = pd.read_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\ETANOL ANP.xlsx')

for sheet_name in planilha_principal.sheetnames:
    worksheet = planilha_principal[sheet_name]
    
    for col_num in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.font = Font(bold=True)
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

planilha_principal.save("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\ETANOL\\ETANOL ANP.xlsx")