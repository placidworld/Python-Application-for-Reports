#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jul 20 14:35:17 2020

@author: AWSCLOUD\l6oi
"""
import pandas
import re

        
def compare_excel_6_2(infile1, infile2, outfile, sk, sheet=None):
    
    print("Comparing Excel 6-2 file ...")
    
    excel1 = pandas.read_excel(infile1, sheet_name=sheet, convert_float = True, header=None, index_col=None, dtype="object")
    excel2 = pandas.read_excel(infile2, sheet_name=sheet, convert_float = True, header=None, index_col=None, dtype="object")


        
    num_pattern = re.compile("^[-0-9.]+$")
    currency_pattern = re.compile("[-0-9,$]+\.[0-9][0-9]?$")
    
    with pandas.ExcelWriter(outfile) as writer:
        work_book = writer.book
        left_align_format = work_book.add_format()
        left_align_format.set_align("left")
        red_format = work_book.add_format()
        red_format.set_font_color('red')
        wrap_format = work_book.add_format({'bold': True})
        wrap_format.set_text_wrap()        
        bold_format = work_book.add_format({'bold': True})

    
        for name1, sheet1 in excel1.items():
            if (name1 not in excel2):
                continue
            
            diff_sheet = work_book.add_worksheet(name1)

            sheet2 = excel2[name1]
            
            sheet1.dropna(1, how = 'all', inplace = True)
            sheet2.dropna(1, how = 'all', inplace = True)
            
            rows = range(0, min(len(sheet1.index), len(sheet2.index)))
            cols = range(0, min(len(sheet1.columns), len(sheet2.columns)))
           

            for row in rows:
                for col in cols:
                    a = sheet1.iloc[row, col]
                    b = sheet2.iloc[row, col]
                    
                    if col >= 2:
                        col -= 1
                        
                    str1 = str(a)
                    str2 = str(b)
                    
                    if str1 == "nan" and str2 == "nan":
                        continue
                    elif str1 == "nan":
                        diff_sheet.write(row, col, "None => {}".format(b))
                        continue
                    elif str2 == "nan":
                        diff_sheet.write(row, col, "{} => None".format(a))
                        continue
    
                    if isinstance(a, int):
                        f1 = "{:,.0f}"
                    elif isinstance(a, str):
                        f1 = "{}"
                    else:
                        f1 = "{:,g}"
                        
                    if isinstance(b, int):
                        f2 = "{:,.0f}"
                    elif isinstance(b, str):
                        f2 = "{}"
                    else:
                        f2 = "{:,g}"
                    
                    try:
                        if currency_pattern.match(str1):
                            f1 = "${:,.2f}"
                            
                        if currency_pattern.match(str2):
                            f2 = "${:,.2f}"                          
                            
    
                        str1 = str1.replace("$", "").replace(",", "").replace("%", "")
                        str2 = str2.replace("$", "").replace(",", "").replace("%", "")
                            
                        if not num_pattern.match(str1) or not num_pattern.match(str2):
                            if str1 != str2:
                                diff_sheet.write(row, col, "{} => {}".format(a, b))
                            else:
                                diff_sheet.write(row, col, str1)
                            continue

                        val1 = float(str1)
                        val2 = float(str2)
                        
                        if str1 == str2:
#                            sheet2.at[row, col] = (f1 + " => " + f2).format(a, b)
                            diff_sheet.write(row, col, '0.00%')
                            continue
                       
                        if (val1 == 0):
                            diff_sheet.write(row, col, (f1 + " => " + f2).format(val1, val2))
                        else:
 #                           sheet2.at[row, col] = (f1 + " => " + f2 + " ({:,.2%})").format(val1, val2, (val2 - val1) / val1)
                            ratio = (val1 - val2)/val1
                            if ratio < -0.5:
                                diff_sheet.write(row, col, "{:,.2%}".format(ratio), red_format)
                            elif ratio > 0.5:
                                diff_sheet.write(row, col, "{:,.2%}".format(ratio), red_format)                                
                            else:
                                diff_sheet.write(row, col, "{:,.2%}".format(ratio))
 #                           sheet2.at[row, col] = "{:,.2%}".format((val2 - val1) / val1)
                              
                    except ValueError:
                        diff_sheet.write(row, col, "{} => {}".format(a, b))
              
            diff_sheet.write(2,1,'')
            h = sheet1.iloc[0, 1] + "\nCurrent SK: " + str(sk.get_csk())
            diff_sheet.merge_range(0, 0, 0, 5, h, wrap_format)
            diff_sheet.set_row(0, 90)
            diff_sheet.set_row(2, 20, bold_format)
            diff_sheet.set_row(3, 20, bold_format)
            diff_sheet.merge_range(2, 0, 3, 0, "Month of Payment", wrap_format)

            diff_sheet.set_column(0, 0, 30, left_align_format)
            diff_sheet.set_column(1, 1, 15, left_align_format)
            diff_sheet.set_column(2, len(sheet1.columns) - 1, 30, left_align_format)
#            diff_sheet.conditional_format(0, 0, len(sheet2.index)-1, len(sheet2.columns)-1,
#                                          {'type': 'cell',
#                                          'criteria': 'less than',
#                                          'value': -0.25,
#                                          'format': red_format}) 
