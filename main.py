# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


import pandas as pd
import glob
import os
import openpyxl
from pandas.io.excel import ExcelWriter
import datetime
import class_currency as basic
import st_Payeer
import st_Whitebit
import Nexo

os.chdir(r"C:\Users\SergeyMalchenko\Desktop\python_files")
if (os.path.isfile("combined_csv.csv")):
  print('Need to delete combined_csv file')
  raise FileExistsError
extension = 'csv'
all_filenames = [i for i in glob.glob('*.{}'.format(extension))]
# combine all files in the list
combined_csv = pd.concat([pd.read_csv(f, sep='\t') for f in all_filenames])

# export to csv
combined_csv = combined_csv.to_csv("combined_csv.csv", index=False, encoding='utf-8-sig')

df1 = pd.read_csv('combined_csv.csv')
df1 = df1['ID транзакции;Номер заказа торговца;Дата операци;Название торговца;Название магазина торговца;Тип операции;Валюта;Стоимость операции;Идентификатор провайдера;Название держателя карты;Маска номера карты;Название страны-эмитента;Название эмитента;Название платежной системы;Платежная система платежной карты;Тип платежной карты'].str.split(';', expand=True)
df1.columns = ['ID транзакции', 'Номер заказа торговца', 'Дата операци', 'Название торговца', 'Название магазина торговца', 'Тип операции', 'Валюта', 'Стоимость операции', 'Идентификатор провайдера', 'Название держателя карты', 'Маска номера карты', 'Название страны-эмитента', 'Название эмитента', 'Название платежной системы', 'Платежная система платежной карты', 'Тип платежной карты']
df1['Стоимость операции'] = df1['Стоимость операции'].astype(float)



# Задаем вид датафрейма и сортируем по мерчу
merch_name = 'Whitebit'                                #REPLACE MERCH NAME
date_from = '2022-10-26'                               # DATE
date_to = '2023-03-13'                                 # DATE
date_from_obj = datetime.datetime.strptime(date_from, '%Y-%m-%d')
date_to_obj = datetime.datetime.strptime(date_to, '%Y-%m-%d')
now_date = datetime.datetime.now().strftime('%Y-%m-%d')
df1 = df1[df1['Название торговца'] == merch_name]  # merchant dataframe



# initiate variables
complete = ['complete']
refund = ['refund']
currency_list = ['EUR', 'USD', 'RUR', 'GBP']  # currencies         ADD ALL CURRENCIES!!!!!!!!!!!!!!!!!!!!!!



# turnover's df & refund's df
df_turnover = df1[df1['Тип операции'].isin(complete)]  # complete transactions
df_turnover_refund = df1[df1['Тип операции'].isin(refund)]  # refund transactions
# Внешний вид фрейма
df1['Идентификатор провайдера'] = ''
df1['Название платежной системы'] = ''
df1.rename(columns={'Идентификатор провайдера': '_', 'Название платежной системы': '_'}, inplace=True)

#  запись Эксель файл

with ExcelWriter(merch_name + '.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:      #REPLACE MERCH NAME
  df1.to_excel(writer, sheet_name='Data') # upload dataframe


statement = merch_name + '.xlsx'   # Name of Excel file (rename)                                           #REPLACE MERCH NAME
book = openpyxl.load_workbook(filename=statement)
statement_sheet = book['Worksheet']
check_sheet = book['Check']

# chose the statement scenario
match merch_name:
  case "Payeer":
    for i in currency_list:
      st_Payeer.payeer_excel(i, df_turnover, df_turnover_refund, statement_sheet, check_sheet, date_from_obj, date_to_obj)  # call excel writer function (st_Payeer file)

  case 'Whitebit':
    for i in currency_list:
      st_Whitebit.whitebit_excel(i, df_turnover, df_turnover_refund, statement_sheet, check_sheet, date_from_obj, date_to_obj)
    basic.wb_holdback(statement_sheet, date_from_obj, date_to_obj, now_date, time=180)

  case 'Nexo Main':
    for i in currency_list:
      Nexo.nexo_excel(i, df_turnover, df_turnover_refund, statement_sheet, check_sheet, date_from_obj, date_to_obj)






book.save(statement)
os.remove("combined_csv.csv")

