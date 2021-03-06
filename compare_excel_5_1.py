# -*- coding: utf-8 -*-
"""
Created on Sat Oct 10 17:03:57 2020

@author: heart
"""

import pandas
import re

        
def compare_excel_5_1(infile1, infile2, outfile, sheet=None, r51data=None):
    
    print("Comparing whole file ...")
    
    ### use pandas to read_excel read excel
    excel1 = pandas.read_excel(infile1, 
                               sheet_name=sheet,
                               convert_float = True, 
                               nrows = 2,
                               skiprows=[0, 1], 
                               index_col=None,
                               header=None,
                               dtype="object")
    excel2 = pandas.read_excel(infile2, 
                               sheet_name=sheet, 
                               convert_float = True, 
                               nrows = 2,
                               skiprows=[0, 1], 
                               index_col=None, 
                               header=None,
                               dtype="object")

    ### number pattern and currency pattern
    num_pattern = re.compile("^[-0-9.]+$")
    currency_pattern = re.compile("[-0-9,$]+\.[0-9][0-9]?$")

    ### use pandas ExcelWriter to create output file    
    with pandas.ExcelWriter(outfile) as writer:
        work_book = writer.book
        left_align_format = work_book.add_format()
        left_align_format.set_align("left")
        red_format = work_book.add_format()
        red_format.set_font_color('red')
        wrap_format = work_book.add_format()
        wrap_format.set_text_wrap()
    
        ### Check the file sheets to make sure compare to the same naming sheets
        for name1, sheet1 in excel1.items():
            if (name1 not in excel2):
                continue
            
            diff_sheet = work_book.add_worksheet(name1)
            diff_sheet.merge_range(0, 0, 0, len(sheet1.columns) - 1, r51data.get_header(), wrap_format)

            sheet2 = excel2[name1]

            ### Remove missing values, how = 'all' 
            """
            axis{0 or ‘index’, 1 or ‘columns’}, default 0
            Determine if rows or columns which contain missing values are removed.

            0, or ‘index’ : Drop rows which contain missing values.
            1, or ‘columns’ : Drop columns which contain missing value.

            Changed in version 1.0.0: Pass tuple or list to drop on multiple axes. Only a single axis is allowed.
            """
            # ‘any’ : If any NA values are present, drop that row or column.
            # ‘all’ : If all values are NA, drop that row or column.
            # Here drop any columns with missing values ALL
            sheet1.dropna(1, how = 'all', inplace = True)
            sheet2.dropna(1, how = 'all', inplace = True)
           
            ### Check # of rows and cols in the file 
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
                        diff_sheet.write(row + 1, col, "None => {}".format(b))
                        continue
                    elif str2 == "nan":
                        diff_sheet.write(row + 1, col, "{} => None".format(a)) # Row starts with 0 by index. 
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
                                diff_sheet.write(row + 1, col, "{} => {}".format(a, b))
                            else:
                                diff_sheet.write(row + 1, col, str1)
                            continue

                        val1 = float(str1)
                        val2 = float(str2)
                        
                        if str1 == str2:
#                            sheet2.at[row, col] = (f1 + " => " + f2).format(a, b)
                            diff_sheet.write(row + 1, col, '0.00%')
                            continue
                       
                        if (val1 == 0):
                            diff_sheet.write(row + 1, col, (f1 + " => " + f2).format(val1, val2))
                        else:
 #                           sheet2.at[row, col] = (f1 + " => " + f2 + " ({:,.2%})").format(val1, val2, (val2 - val1) / val1)
                            ratio = (val1 - val2)/val1
                            if ratio < -0.5:
                                diff_sheet.write(row + 1, col, "{:,.2%}".format(ratio), red_format)
                            elif ratio > 0.5:
                                diff_sheet.write(row + 1, col, "{:,.2%}".format(ratio), red_format)                                
                            else:
                                diff_sheet.write(row + 1, col, "{:,.2%}".format(ratio))
 #                           sheet2.at[row, col] = "{:,.2%}".format((val2 - val1) / val1)
                              
                    except ValueError:
                        diff_sheet.write(row + 1, col, "{} => {}".format(a, b))

#            sheet2.dropna(axis = 1, how='all', inplace=True)
#            sheet2.to_excel(writer, sheet_name = name1, header=None, index = False)
#            diff_sheet = writer.sheets[name1]
            diff_sheet.set_column(0, len(sheet1.columns) - 1, 30, left_align_format)
            diff_sheet.set_row(0, 120)
#            diff_sheet.conditional_format(0, 0, len(sheet2.index)-1, len(sheet2.columns)-1,
#                                          {'type': 'cell',
#                                          'criteria': 'less than',
#                                          'value': -0.25,
#                                          'format': red_format}) 
