#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd,glob
from pandas import DataFrame
import pandas as pd
import numpy as np
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

def readFile(file_list) :
    worksheet_list = []
    for f in file_list :
        workbook = xlrd.open_workbook(f)
        temp = {'worksheet' : workbook.sheet_by_index(0), 'nrows' : workbook.sheet_by_index(0).nrows }
        worksheet_list.append(temp)
    return worksheet_list

def isNumber(s) :
    try :
        float(s)
        return True
    except ValueError :
        return False

def makeDataFrame(worksheet_list) :
    row_val = []
    frame_list = []
    date = ""
    for worksheet in worksheet_list :
        for row_num in range(worksheet['nrows']) :
            if str(worksheet['worksheet'].row_values(row_num)[0]).find("년") != -1 :
                date = worksheet['worksheet'].row_values(row_num)[0]
            elif isNumber(worksheet['worksheet'].row_values(row_num)[0]) :
                row_val.append(worksheet['worksheet'].row_values(row_num))
            elif str(worksheet['worksheet'].row_values(row_num)[0]).find("합계") != -1 :
                frame = DataFrame(row_val, columns = ['rank','name','start_date','sales_account','sales_account_ratio','sales_account_varience','sales_account_varience_ratio','accum_sales_account','audience','audience_varience','audience_varience_ratio','accum_audience', 'screen','show','main_country','country','maker','distributor','rating','genre','director','actor'])
                del frame['sales_account_ratio']
                del frame['sales_account_varience']
                del frame['sales_account_varience_ratio']
                del frame['accum_sales_account']
                del frame['screen']
                frame['date'] = date
                frame.fillna(0)
                frame_list.append(frame)
                row_val = []
                date = ""
    return frame_list

def writeDataFrame(frame_list) :
    f = open("frame_data.txt",'w')
    f.close()
    for frame in frame_list :
        frame.to_csv("frame_data.txt", sep = '\t', mode = 'a', encoding = 'utf-8', index = False)
        with open("frame_data.txt",'a') as f:
            f.write("\n")
    print "success"

file_list = glob.glob('*.xlsx')

worksheet_list = readFile(file_list)

frame_list = makeDataFrame(worksheet_list)

print "All has been done!"
