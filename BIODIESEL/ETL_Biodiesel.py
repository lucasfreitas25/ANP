from urllib.request import urlretrieve 
from zipfile import ZipFile
import pandas as pd 
import openpyxl
from openpyxl.styles import Border, Side, Font
from openpyxl.utils import get_column_letter


url = 'https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/arquivos/pb-da-biodiesel.zip'
arquivo = 'pb-da-biodiesel.zip'

urlretrieve(url, arquivo)

with ZipFile(arquivo, "r") as zip:
    zip.extractall('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados em csv')
    
df_capacidade = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados em csv\\Biodiesel_DadosAbertos_CSV_Capacidade.csv')
df_capacidade['CNPJ'] = df_capacidade['CNPJ'].astype(str)
df_capacidade.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados Biodiesel Capacidade.xlsx', index=False)

df_matprima = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados em csv\\Biodiesel_DadosAbertos_CSV_MatériaPrima.csv')
df_matprima['Quantidade (m³)'] = df_matprima['Quantidade (m³)'].str.replace(',', '.').astype(float)
df_matprima.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados Biodiesel Materia Prima.xlsx', index=False)

df_prod = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados em csv\\Biodiesel_DadosAbertos_CSV_Produç╞o.csv')
df_prod['Produção de Biodiesel'] = df_prod['Produção de Biodiesel'].str.replace(',', '.').astype(float)
df_prod.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados Biodiesel Produção.xlsx', index=False)

df_vendas = pd.read_csv('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados em csv\\Biodiesel_DadosAbertos_CSV_Vendas.csv')
df_vendas['Vendas de Biodiesel'] = df_vendas['Vendas de Biodiesel'].str.replace(',', '.').astype(float)
df_vendas.to_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados Biodiesel Vendas.xlsx', index=False)


planilha_principal = openpyxl.Workbook()

wb_capacidade = openpyxl.load_workbook('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados Biodiesel Capacidade.xlsx')  
wb_matprima = openpyxl.load_workbook("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados Biodiesel Materia Prima.xlsx")  
wb_prod = openpyxl.load_workbook('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados Biodiesel Produção.xlsx')
wb_vendas = openpyxl.load_workbook('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\Dados Biodiesel Vendas.xlsx')


aba_capacidade = planilha_principal.create_sheet("CAPACIDADE")
aba_matprima = planilha_principal.create_sheet("MATÉRIA PRIMA")
aba_prod = planilha_principal.create_sheet("PRODUÇÃO")
aba_vendas = planilha_principal.create_sheet("VENDAS")

for linha in wb_capacidade.active.iter_rows(values_only=True):
    aba_capacidade.append(linha)

for linha in wb_matprima.active.iter_rows(values_only=True):
    aba_matprima.append(linha)
    
for linha in wb_prod.active.iter_rows(values_only=True):
    aba_prod.append(linha)
    
for linha in wb_vendas.active.iter_rows(values_only=True):
    aba_vendas.append(linha)

for aba in planilha_principal.sheetnames:
    if aba not in ["CAPACIDADE", "MATÉRIA PRIMA", "PRODUÇÃO", "VENDAS"]:
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

lista_abas = [aba_capacidade, aba_matprima, aba_prod, aba_vendas]
for abas in lista_abas:
    ajustar_colunas(abas)
    
    
############################    
worksheet = planilha_principal.active
df = pd.read_excel('C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\BIODIESEL ANP.xlsx')

for sheet_name in planilha_principal.sheetnames:
    worksheet = planilha_principal[sheet_name]
    
    for col_num in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.font = Font(bold=True)
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
planilha_principal.save("C:\\Users\\LucasFreitas\\Documents\\Lucas Freitas Arquivos\\DATAHUB\\DADOS\\ANP\\BIODIESEL\\BIODIESEL ANP.xlsx")