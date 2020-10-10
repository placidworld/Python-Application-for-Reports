# -*- coding: utf-8 -*-
"""
Created on Sat Oct 10 16:26:40 2020

@author: heart
"""



import sys
import pandas
import numpy
import os
import re
import ibm_db
import sqlalchemy
import datetime

from . import compare_5_1_with_db2

from IDDOC.sys_connect import hostname, username, password, database, IDR_host, path, outpath

pandas.set_option('display.float_format', lambda x: '%.3f' % x)


def read_51_all(writer, infile1, infile2, file_number):    
    print("Comparing " + infile1 + " with " + infile2 + " ...")
    
    workbook = writer.book

    count_format = workbook.add_format({"num_format": "#,##0"})

    merge_format = workbook.add_format({'align': 'left'})

    wrap_format = workbook.add_format()
    wrap_format.set_text_wrap()

    pct_format = workbook.add_format({"num_format": "0.00%"})

    excel1 = pandas.read_excel(path + infile1, 
                               sheet_name = None, 
                               convert_float = True, 
                               header = 2, 
                               index_col = None, 
                               nrows = 1,
                               dtype="object")

    excel2 = pandas.read_excel(path + infile2, 
                               sheet_name = None,
                               convert_float = True, 
                               header = 2, 
                               index_col = None, 
                               nrows = 1,
                               dtype="object")
                
    pct_pat = re.compile(r'^(.*)%')

    for name in excel1:
        if name not in excel2:
            continue

        sheet1 = excel1[name]
        sheet2 = excel2[name]
        
        # remove Unnamed columns from both Excel1 and Excel2
        dfcols1 = list(sheet1.columns)

        for c in sheet1.columns:
            if isinstance(c, str) and "Unnamed" in c:
                dfcols1.remove(c)

        sheet1 = sheet1.loc[:, dfcols1]
        
        dfcols2 = list(sheet2.columns)

        for c in sheet2.columns:
            if isinstance(c, str) and "Unnamed" in c:
                dfcols2.remove(c)

        sheet2 = sheet2.loc[:, dfcols2]
        
        if file_number == 0:
            out_sheet = writer.book.add_worksheet(name)
            writer.sheets[name] = out_sheet
        else:
            out_sheet = writer.sheets[name]

        if name == "Total Statistics":
            out_sheet.set_column("A:A", 20)
            out_sheet.set_column(1, 1, 30, pct_format)
            out_sheet.set_column(2, 2, 60, pct_format)
            out_sheet.set_column(3, 3, 60, pct_format)
            out_sheet.set_column(4, 4, 30, pct_format)
            out_sheet.set_column(5, 5, 30, pct_format)
            out_sheet.set_column(6, 6, 40, pct_format)
            out_sheet.set_column(7, 7, 40, pct_format)

        elif name == "Substance Abuse":
            out_sheet.set_column("A:A", 20)
            out_sheet.set_column(1, 1, 30, pct_format)
            out_sheet.set_column(2, 2, 30, pct_format)
            out_sheet.set_column(3, 3, 30, pct_format)
            out_sheet.set_column(4, 4, 30, pct_format)
            out_sheet.set_column(5, 5, 40, pct_format)
            out_sheet.set_column(6, 6, 50, pct_format)
            out_sheet.set_column(7, 7, 50, pct_format)
            out_sheet.set_column(8, 8, 30, pct_format)
            out_sheet.set_column(9, 9, 30, pct_format)
            out_sheet.set_column(10, 10, 30, pct_format)
            out_sheet.set_column(11, 11, 30, pct_format)
            out_sheet.set_column(12, 12, 30, pct_format)
            out_sheet.set_column(13, 13, 30, pct_format)
            out_sheet.set_column(14, 14, 30, pct_format)

        elif name == "PBP AIPBP":
            out_sheet.set_column("A:A", 20)
            out_sheet.set_column(1, 1, 40, pct_format)
            out_sheet.set_column(2, 2, 40, pct_format)
            out_sheet.set_column(3, 3, 40, pct_format)
            out_sheet.set_column(4, 4, 40, pct_format)
            out_sheet.set_column(5, 5, 40, pct_format)
            out_sheet.set_column(6, 6, 40, pct_format)

        elif name == "Exclusions":
            out_sheet.set_column("A:A", 20)
            out_sheet.set_column(1, 1, 40, pct_format)
            out_sheet.set_column(2, 2, 40, pct_format)
            out_sheet.set_column(3, 3, 50, pct_format)
            out_sheet.set_column(4, 4, 40, pct_format)
            out_sheet.set_column(5, 5, 40, pct_format)
            out_sheet.set_column(6, 6, 30, pct_format)
            out_sheet.set_column(7, 7, 80, pct_format)
            out_sheet.set_column(8, 8, 60, pct_format)
            out_sheet.set_column(9, 9, 90, pct_format)
            out_sheet.set_column(10, 10, 30, pct_format)
            out_sheet.set_column(11, 11, 30, pct_format)                    

        for col in range(0,len(sheet1.columns)):
            if file_number == 0:
                out_sheet.write(0, col, sheet1.columns[col])
                
            if col == 0 :
                out_sheet.write(file_number + 1, 0, sheet1.iloc[0, col])                 
            else:
                v1 = sheet1.iloc[0, col]
                v2 = sheet2.iloc[0, col]                

                if str(v1) == "nan" or str(v2) == "nan":
                    out_sheet.write(file_number + 1, col, f"{v1} => {v2}")
                    continue
                
                if isinstance(v1, str):
                    m = pct_pat.match(v1)
                    if m:
                        v1 = float(m.group(1)) / 100

                if isinstance(v2, str):
                    m = pct_pat.match(v2)
                    if m:
                        v2 = float(m.group(1)) / 100                    

                if v1 == 0:
                    v = f"0 => {v2}"
                    out_sheet.write(file_number + 1, col, v)
                else:
                    try:
                        v = (v2 - v1) / v1
                        out_sheet.write(file_number + 1, col, "{:,.2%}".format(v))

                    except ValueError as e:
                        print(f"exception1: {e}")

 #   excel1.iloc[0,]

def do_process(files_prev, files_curr, outfile, model, report_num):
    with pandas.ExcelWriter(outpath + outfile) as writer:

        n = 0

        for ppo in files_prev:
            if ppo in files_curr:
                read_51_all(writer, files_prev[ppo], files_curr[ppo], n)
                n += 1

def PROGRAM1_5_1_ppo_Sum(previous_mon=None, current_mon=None):
    model = "PROGRAM1"
    report_num = "5_1"

    if current_mon is None:
        current_mon = input("Enter report month current (YYYYMM): ")
        previous_mon = input("Enter report month previous (YYYYMM): ") 

    ppo_name = "[A-Za-z0-9]+"

    outfile = model + "_" + report_num + "_ppo_" + current_mon + "_Validation.xlsx"
    
    files_prev = {}
    files_curr = {}

    tag = None
    
    pattern = re.compile("^" + model + "_" + report_num + "_ ?(" + ppo_name + ")_")
    
    files = os.listdir(path)

    files.sort()

    for fname in files:
        if "Validation" in fname:
            continue
        
        if "Summary" in fname:
            continue

        if "Differences" in fname:
            continue

        m = pattern.match(fname)

        if not m:
            continue       

        print("matched:" + fname)

        tag = m.group(1)
        
        if previous_mon in fname:
            files_prev[tag] = fname            

        elif current_mon in fname:
            files_curr[tag] = fname

    do_process(files_prev, files_curr, outfile, model, report_num)


if __name__ == '__main__':
    PROGRAM1_5_1_ppo_Sum()