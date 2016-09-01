import xlrd,glob
from pandas import DataFrame
import pandas as pd
import numpy as np

file_list = glob.glob('*.xlsx')
file_list.sort()

worksheet_list = []

for f in file_list :
    workbook = xlrd.open_workbook(f)
    temp = {'date' : f, 'worksheet' : workbook.sheet_by_index(0), 'nrows' : workbook.sheet_by_index(0).nrows }
    worksheet_list.append(temp)

def isNumber(s) :
    try :
        float(s)
        return True
    except ValueError :
        return False

row_val_list = []

for worksheet in worksheet_list :
    row_val = []
    for row_num in range(worksheet['nrows']) :
        if isNumber(worksheet['worksheet'].row_values(row_num)[0]) :
            row_val.append(worksheet['worksheet'].row_values(row_num))
    temp = {'date' : worksheet['date'], 'row_val' : row_val}
    row_val_list.append(temp)

frame_list = []

for obj in row_val_list :
    frame = DataFrame(obj['row_val'],columns = ['rank','name','start_date','sales_account','sales_account_ratio','sales_account_varience','sales_account_varience_ratio','accum_sales_account','audience','audience_varience','audience_varience_ratio','accum_audience', 'screen','show','main_country','country','maker','distributor','rating','genre','director','actor'])
    del frame['sales_account']
    del frame['sales_account_ratio']
    del frame['sales_account_varience']
    del frame['sales_account_varience_ratio']
    del frame['accum_sales_account']
    del frame['screen']
    del frame['country']
    del frame['main_country']
    del frame['maker']
    del frame['distributor']
    del frame['rating']
    del frame['genre']
    
    frame.fillna(0)
    temp = {'date' : obj['date'], 'frame' : frame}
    frame_list.append(temp)

f = open("frame_data.txt",'w')
f.close()

for obj in frame_list :
    with open("frame_data.txt",'a') as f :
        f.write(obj['date'] + "\n")
    obj['frame'].to_csv("frame_data.txt", sep = '\t', mode = 'a', encoding = 'utf-8', index = False)
    with open("frame_data.txt",'a') as f :
        f.write("\n")
    print obj['date']

print "success"
