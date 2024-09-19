import pandas as pd
import re

file_path = 'G:\\ESCRITA FISCAL\\COLABORADORES\\ANDRÉ FELIPE\\208\\PLANILHA 208.xlsx'

xls = pd.ExcelFile(file_path)

sheet_name = input("Digite o nome da planilha que deseja usar: ").upper()

df = pd.read_excel(file_path, sheet_name=sheet_name)

def remove_pontuacao(text):
   
    return re.sub(r'[^\w\s]', '', str(text))


df['LOCATARIO'] = df['LOCATARIO'].apply(remove_pontuacao)

df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce', format='%d/%m/%Y')

if df['DATA'].isnull().any():
    print("Atenção: Algumas datas não foram convertidas corretamente.")
    
df['DATA'] = df['DATA'].dt.strftime('%d%m%Y')

locatarios = df['LOCATARIO']
datas = df['DATA']
valores = df['VALOR']
piss = df['PIS']
cofinss = df['COFINS']

espaço = 20 * ' '
espaço2 = 520 * ' '
espaço3 = 55 * ' ' 
def format_data(locatario, data, valor, pis, cofins):

    return f"1{f'{locatario}'.zfill(14)}{espaço}{data}{f'{valor:.2f}'.zfill(14)}01{f'{valor:.2f}'.zfill(14)}000.6500{f'{pis:.2f}'.zfill(14)}01{f'{valor:.2f}'.zfill(14)}003.00{f'{cofins:.2f}'.zfill(14)}  000{espaço3}61300"

for locatario, data, valor, pis, cofins in zip(locatarios, datas, valores, piss, cofinss):
    formatted_data = format_data(locatario, data, valor, pis, cofins)
    

output_file = 'resultado_formatado.txt'

with open(output_file, 'w') as f:
    
    for locatario, data, valor, pis, cofins in zip(locatarios, datas, valores, piss, cofinss):
        formatted_data = format_data(locatario, data, valor, pis, cofins)
        f.write(formatted_data + '\n')  
