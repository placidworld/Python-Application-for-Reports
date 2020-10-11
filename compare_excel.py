# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 13:57:40 2020

@author: heart
"""


import pandas
import re

        
def compare_excel(infile1, infile2, outfile, sheet=None):
    
    print("Comparing whole file ...")
    
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
                              
                    except ValueError as e:
                        diff_sheet.write(row, col, "{} => {}".format(a, b))

#            sheet2.dropna(axis = 1, how='all', inplace=True)
#            sheet2.to_excel(writer, sheet_name = name1, header=None, index = False)
#            diff_sheet = writer.sheets[name1]
            diff_sheet.set_column(0, len(sheet1.columns) - 1, 30, left_align_format)
#            diff_sheet.conditional_format(0, 0, len(sheet2.index)-1, len(sheet2.columns)-1,
#                                          {'type': 'cell',
#                                          'criteria': 'less than',
#                                          'value': -0.25,
#                                          'format': red_format}) 
