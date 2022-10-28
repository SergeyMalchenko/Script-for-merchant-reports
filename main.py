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
import Merchant1
import Merchant2
import Merchant3

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
df1 = df1['***'].str.split(';', expand=True)
df1.columns = ['***']
df1['С***'] = df1['***'].astype(float)



# Задаем вид датафрейма и сортируем по мерчу
merch_name = '***'                                #REPLACE MERCH NAME
date_from = '***'                               # DATE
date_to = '***'                                 # DATE
date_from_obj = datetime.datetime.strptime(date_from, '%Y-%m-%d')
date_to_obj = datetime.datetime.strptime(date_to, '%Y-%m-%d')
now_date = datetime.datetime.now().strftime('%Y-%m-%d')
df1 = df1[df1['***'] == merch_name]  # merchant dataframe



# initiate variables
complete = ['complete']
refund = ['refund']
currency_list = ['EUR', 'USD', 'RUR', 'GBP' ***]  # currencies         



# turnover's df & refund's df
df_turnover = df1[df1['**'].isin(***)]  # complete transactions
df_turnover_refund = df1[df1['***'].isin(***)]  # refund transactions
# Внешний вид фрейма
df1['***'] = ''
df1['***'] = ''
df1.rename(columns={'***': '_', '***': '_'}, inplace=True)

#  запись Эксель файл

with ExcelWriter(merch_name + '.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:      
  df1.to_excel(writer, sheet_name='Data') # upload dataframe


statement = merch_name + '.xlsx'   # Name of Excel file                                           
book = openpyxl.load_workbook(filename=statement)
statement_sheet = book['Worksheet']
check_sheet = book['Check']

# chose the statement scenario
match merch_name:
  case "Merchant1":
    for i in currency_list:
      merchant1.merchant1_excel(i, df_turnover, df_turnover_refund, statement_sheet, check_sheet, date_from_obj, date_to_obj)  # call excel writer function 

Merchant12    for i in currency_list:
      Merchant2.Merchant2_excel(i, df_turnover, df_turnover_refund, statement_sheet, check_sheet, date_from_obj, date_to_obj)
    basic.wb_holdback(statement_sheet, date_from_obj, date_to_obj, now_date, time=180)

  case 'Merchant3':
    for i in currency_list:
      Merchant3.Merchant3_excel(i, df_turnover, df_turnover_refund, statement_sheet, check_sheet, date_from_obj, date_to_obj)






book.save(statement)
os.remove("combined_csv.csv")

