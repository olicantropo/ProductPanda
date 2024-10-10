import pandas as pd
import re

arquivo_a = pd.read_excel('myfile.xlsx')
arquivo_b = pd.read_csv('Wix_Templates_Products_CSV.csv')

colunas_desejadas = [
    'name', 'sku', 'inventory', 'collection', 'price', 'ribbon', 
    'fieldType', 'handleId', 'productOptionType1', 'productOptionName1', 
    'productOptionDescription1', 'Visible'
]

df_adaptado = arquivo_a[colunas_desejadas]

df_adaptado['sku'] = df_adaptado['sku'].fillna(0).astype(int) 
df_adaptado['inventory'] = df_adaptado['inventory'].fillna(0).astype(int)

df_adaptado['price'] = df_adaptado['price'].fillna(0).round(2)  

df_adaptado['Visible'] = df_adaptado['Visible'].fillna('').astype(str).str.upper()

def preencher_name(grupo):
    primeiro_name_valido = grupo['name'].loc[grupo['name'].str.strip() != ''].iloc[0]
    grupo['name'] = grupo['name'].replace('', primeiro_name_valido).replace(' ', primeiro_name_valido)
    return grupo

df_adaptado = df_adaptado.groupby('handleId').apply(preencher_name)

def ajustar_letras_avulsas(descricao):
    if pd.isna(descricao):
        return descricao
    
    return re.sub(r'\b[a-zA-Z]\b', lambda x: x.group().upper(), descricao)

df_adaptado['productOptionDescription1'] = df_adaptado['productOptionDescription1'].apply(ajustar_letras_avulsas)

def extrair_numero_handleId(handleId):
    match = re.search(r'\d+', handleId)
    return int(match.group()) if match else 0 

df_adaptado['handleId_num'] = df_adaptado['handleId'].apply(extrair_numero_handleId)

df_adaptado = df_adaptado.sort_values(by=['handleId_num', 'fieldType'], ascending=[True, True])

df_adaptado = df_adaptado.drop(columns=['handleId_num'])

df_adaptado.to_csv('adapt.csv', index=False)

print("File adapted to Wix Stores.")
