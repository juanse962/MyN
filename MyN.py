import pandas as pd
import re 
import numpy as np


path = str(input("Nombre del archivo de excel SIN extension xlsx: ") + '.xlsx')
xls = pd.ExcelFile(path)
df = xls.parse(xls.sheet_names[0])
df.columns = ['CANTIDAD/TIPO','TIENDA/CONTACTO','VALOR COMPRA','VALOR VENTA','VALOR COMISIÓN','GANANCIA']
clients = int(input("Numero de clientes que tienes: "))
def quantity(df,clients):

    countA = 0.0
    countB = 0.0
    countAA = 0.0
    number = 0.0
    l = []
    l = df['CANTIDAD/TIPO'][5::]
    clients += 5
    for i in range(5,clients):
        if 'AA' in str(l[i]):
            number = l[i].split('AA')
            countAA += float(number[0])
            continue
        elif 'B' in str(l[i]):
            number = l[i].split('B')
            countB += float(number[0])
            continue
        elif 'A' in str(l[i]):
            number = l[i].split('A')
            countA += float(number[0])
            continue
    
    return countA,countB,countAA

quantity(df,clients)
def buy_value(df,clients):
    buy_value = 0
    clients += 5
    for i in range(5,clients):
        df['VALOR COMPRA'][i] =  str(df['VALOR COMPRA'][i])
        if  'nan' in df['VALOR COMPRA'][i]:
            continue
        buy_value += int(df['VALOR COMPRA'][i])
    return buy_value

def commission_value(df,clients):
    commission_value = 0
    clients += 5
    for i in range(5,clients):
        df['VALOR COMISIÓN'][i] =  str(df['VALOR COMISIÓN'][i])
        if  'nan' in df['VALOR COMISIÓN'][i]:
            continue
        commission_value += int(df['VALOR COMISIÓN'][i])
    return commission_value

def sale_value(df,clients):
    sale_value = 0
    clients += 5
    for i in range(5,clients):
        df['VALOR VENTA'][i] =  str(df['VALOR VENTA'][i])
        if  'nan' in df['VALOR VENTA'][i]:
            continue
        sale_value += int(df['VALOR VENTA'][i])
    return sale_value

buy_values = buy_value(df,clients)
sale_value = sale_value(df,clients)
commission_value = commission_value(df,clients)

def stores(df):
    
    array = []
    for i in range(5,df['TIENDA/CONTACTO'].size):
        
        if  'nan' in str(df['TIENDA/CONTACTO'][i]):
            continue
        array.append(str(df['TIENDA/CONTACTO'][i]))
    return array

def gain(df,clients):

    array = []
    clients += 6
    for i in range(5,df['VALOR VENTA'].size):
        df['VALOR COMPRA'][i] =  str(df['VALOR COMPRA'][i])
        df['VALOR VENTA'][i]  =  str(df['VALOR VENTA'][i])
        df['VALOR COMISIÓN'][i] =  str(df['VALOR COMISIÓN'][i])

        if  'nan' in df['VALOR VENTA'][i]:
            continue
            
        if 'nan' in df['VALOR COMPRA'][i]:
            continue
            
        if 'nan' in df['VALOR COMISIÓN'][i]:
            continue
        else:
            array.append('$'+ str(int(df['VALOR VENTA'][i]) - int(df['VALOR COMPRA'][i]) - int(df['VALOR COMISIÓN'][i])))
    return array

a = stores(df)
arr = gain(df,clients)
def value_gain(arr):
    aux = list(arr)
    count = 0
    for i in range(len(aux)):
        aux[i] = aux[i].replace('$','')
        count += int(aux[i])
    return count
l = []
value = value_gain(arr)
l.append('$'+str(value))

amountA,amountB,amountAA = quantity(df,clients)
data = {
        'TIPO':                         ['Huevo A', 'Huevo B','Huevo AA'],
        'CANTIDADADES  TOTALES':        [str(amountA)  ,  str(amountB) , str(amountAA) ],
        'CANTIDADAD DE HUEVOS':         [str(amountA*300)  ,  str(amountB*300) , str(amountAA*300) ],
        'VALORES':                      ['VALOR COMPRA', 'VALOR VENTA','VALOR COMISIÓN'],
        'VALORES TOTALES':              [ '$'+str(buy_values), '$' + str(sale_value), '$' + str(commission_value)  ],
}


df = pd.DataFrame (data, columns = ['TIPO','CANTIDADADES  TOTALES','CANTIDADAD DE HUEVOS','VALORES','VALORES TOTALES'])
df2 = pd.DataFrame(a, columns = ['TIENDA/CONTACTO'])
df = pd.concat([df,df2],axis=1)
df2 = pd.DataFrame(arr, columns = ['GANANCIAS'])
df = pd.concat([df,df2],axis=1)
df2 = pd.DataFrame(l, columns = ['GANANCIAS TOTALES'])
df = pd.concat([df,df2],axis=1)
df.to_excel("output.xlsx", sheet_name= 'Resultado')

writer = pd.ExcelWriter("output.xlsx", engine='xlsxwriter')

df.to_excel(writer, sheet_name='Resultado')

workbook  = writer.book
worksheet = writer.sheets['Resultado']

format1 = workbook.add_format({'num_format': '#,##0.00'})
format2 = workbook.add_format({'num_format': '0%'})

worksheet.set_column('B:B', 18)
worksheet.set_column('C:C', 25)
worksheet.set_column('D:D', 25, format2)
worksheet.set_column('E:E', 18, format1)
worksheet.set_column('F:F', 18, format2)
worksheet.set_column('G:G', 50, format1)
worksheet.set_column('H:H', 25, format1)
worksheet.set_column('I:I', 25, format1)
writer.save()