# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 14:47:12 2020

@author: heart
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jul 20 15:08:15 2020

@author: AWSCLOUD\l6oi
"""


import pandas
import re

def compare_phi_6_3(infile1, infile2, summary_file, diff_file, header, index_col, model, report_generated, sk, aco_name, di_6_3, tot_change, null_change):
 
    csk = sk.get_csk()
    psk = sk.get_psk()

    excel1 = pandas.read_excel(infile1, sheet_name=None, header=header, index_col=None)
    excel2 = pandas.read_excel(infile2, sheet_name=None, header=header, index_col=None)

    excel0 = pandas.read_excel(infile2, sheet_name=0, 
                                        header=None, 
                                        index_col=None, 
                                        nrows = 3,
                                        convert_float = True, 
                                        dtype="object")
    
    py = excel0.iloc[2,0]
    
    if model == 'PROGRAM1':
        run_dt = excel0.iloc[1,13]
    else:
        run_dt = excel0.iloc[1,17]        
  
    ### Create excel output    
    diff_writer = pandas.ExcelWriter(diff_file)
    diff_book = diff_writer.book
    red_format = diff_book.add_format({})
    red_format.set_font_color('red')
    
    summary_writer = pandas.ExcelWriter(summary_file)
    summary_book = summary_writer.book
    count_format = summary_book.add_format({"num_format": "#,##0"})
    left_align_format = summary_book.add_format()
    left_align_format.set_align("left")
    pct_format = summary_book.add_format({"num_format": "0.00%"})   
    
    sheet_list_1 = list(excel1)
    sheet_list_2 = list(excel2)
    
    for i in range(0, len(sheet_list_1)):
        
        name1 = sheet_list_1[i]
        name2 = sheet_list_2[i]
        
        sheet1 = excel1[name1]
        sheet2 = excel2[name2]
        
        if sheet1.iloc[0, 0] == "PREF":
            provider_type = "PREF"
        else:
            provider_type = "PART"
            
        sheet1 = sheet1.loc[lambda df: df['Provider Type'] == provider_type, :]
        sheet2 = sheet2.loc[lambda df: df['Provider Type'] == provider_type, :]
        
        if "Physician" in name1:
            index_name = "Individual NPI"
        else:
            index_name = "Organization NPI"
        
        # remove duplicate records in previous month
        previous_records = len(sheet1.index)
        sheet1.drop_duplicates(index_name, inplace=True)
        sheet1.set_index(index_name, inplace=True)
        previous_duplicate_records = previous_records - len(sheet1.index)
        
        # remove duplicate records in current month
        current_records = len(sheet2.index)
        sheet2.drop_duplicates(index_name, inplace=True)
        sheet2.set_index(index_name, inplace=True)
        current_duplicate_records = current_records - len(sheet2.index)
        
        if previous_records == 0:
            if current_records != 0:
                di_6_3.collect_stats(aco_name, name1, f"Current vs Previous > {tot_change}%")
        elif abs(current_records - previous_records) / previous_records > tot_change/100:
            di_6_3.collect_stats(aco_name, name1, f"Current vs Previous > {tot_change}%")
            
            
        # remove Unnamed columns
        columns = list(sheet1.columns)
        for c in sheet1.columns:
            if isinstance(c, str) and "Unnamed" in c:
                columns.remove(c)
                
        sheet1 = sheet1.loc[:, columns]
        sheet2 = sheet2.loc[:, columns]
        
        print("Comparing sheet: " + str(name1) + ", index: " + str(index_name))
                
        merged = pandas.merge(sheet1, sheet2, left_index=True, right_index=True, how="outer")
        
        # remove dupliate keys
        rows = list(dict.fromkeys(merged.index))
        
        result_sheet = pandas.DataFrame(index = rows, columns = columns, dtype = 'object')
        result_sheet.index.name = index_name
        
        both_records = 0
        changed_records = 0
        
        result_sheet.to_excel(diff_writer, sheet_name = name1, index = True)
        s = diff_writer.sheets[name1]
        
        for r in range(0, len(rows)):
            row = rows[r]
            if merged.loc[row, 'Provider Type_x'] == provider_type and merged.loc[row, 'Provider Type_y'] == provider_type:
                both_records += 1
                
            changed = 0
            for c in range(0, len(columns)):
                col = columns[c]
                cell1 = merged.loc[row, col + "_x"]
                cell2 = merged.loc[row, col + "_y"]
                
                str1 = str(cell1)
                str2 = str(cell2)
                
                ### Data cleaning
                if "incurred" in col or "amount" in col:
                    f1 = "${:,.2f}"
                    f2 = "${:,.2f}"
                    if isinstance(cell1, str):
                        cell1 = cell1.replace("$", "").replace(",", "")
                        if "(" in cell1:
                            # negative value
                            cell1 = cell1.replace("(", "").replace(")", "")
                            cell1 = - float(cell1)
                        else:
                            cell1 = float(cell1)
                    if isinstance(cell2, str):
                        cell2 = cell2.replace("$", "").replace(",", "")
                        if "(" in cell2:
                            # negative value
                            cell2 = cell2.replace("(", "").replace(")", "")
                            cell2 = - float(cell2)
                        else:
                            cell2 = float(cell2)
                else:
                    if isinstance(cell1, float):
                        f1 = "{:,.0f}"
                    else:
                        f1 = "{}"
                        
                    if isinstance(cell2, float):
                        f2 = "{:,.0f}"
                    else:
                        f2 = "{}"
                                    
                try:
                    if str1 == 'nan' and str2 == 'nan':
                        #result_sheet.loc[row, col] = ""
                        s.write(r + 1, c + 1, "")
                    
                    elif str1 == 'nan':
                        #result_sheet.loc[row, col] = ("None => " + f2).format(cell2)
                        s.write(r + 1, c + 1, ("None => " + f2).format(cell2))
                        changed = 1
                    
                    elif str2 == 'nan':
                        #result_sheet.loc[row, col] = (f1 + " => None").format(cell1)
                        s.write(r + 1, c + 1, (f1 + " => None").format(cell1))
                        changed = 1
                    
                    elif cell1 != cell2:
                        changed = 1
                        if isinstance(cell1, str) or isinstance(cell2, str) or cell1 == 0:
                            #result_sheet.loc[row, col] = (f1 + " => " + f2).format(cell1, cell2)
                            s.write(r + 1, c + 1, (f1 + " => " + f2).format(cell1, cell2))
                        else:
                            diff = (cell2-cell1)/cell1
                            if diff < -0.5 or diff > 0.5:
                                #result_sheet.loc[row, col] = "{:.2%}".format((cell2-cell1)/cell1)
                                s.write(r + 1, c + 1, "{:.2%}".format(diff), red_format)
                            else:
                                s.write(r + 1, c + 1, "{:.2%}".format(diff))

                    else:
                        #result_sheet.loc[row, col] = f1.format(cell1)
                        s.write(r + 1, c + 1, f1.format(cell1))
                except ValueError as e:
                    print("error: " + str(e))
                    print("sheet: " + name1 + ", NPI: " + str(row) + ", column: " + str(col) + ", cell1: " + str1 + ", cell2: " + str2)
                    #result_sheet.loc[row, col] = f1.format(cell1)
                    break
                
            if changed == 1:
                changed_records += 1
        

        s.set_column(0, 1, 20)
        s.set_column(2, 2, 10)
        s.set_column(3, 3, 50)
        s.set_column(4, 13, 40)


        
        summary = pandas.DataFrame({
                'Type': ['Records in Previous Month',
                         'Records in Current Month',
                         'Records in Both Months',
                         'Records in Previous Month Only',
                         'Records in Current Month Only',
                         'Duplicate Records in Previous Month',
                         'Duplicate Records in Current Month',
                         'Records Changed'],
                'Count': [previous_records,
                          current_records,
                          both_records,
                          previous_records - both_records,
                          current_records - both_records,
                          previous_duplicate_records,
                          current_duplicate_records,
                          changed_records]
                })

 
    
        summary.to_excel(summary_writer, sheet_name = name1, index=False, startrow=1)

        workbook = summary_writer.book
        merge_format = workbook.add_format({'align': 'left'})
        wrap_format = workbook.add_format()
        wrap_format.set_text_wrap()

        header = f"Report 6-3\nSK used:  {csk}\n{py}\nRun date: {run_dt}"
        di_6_3.set_header(header)

        summary_sheet = summary_writer.sheets[name1]
        summary_sheet.merge_range('A1:B1', header, wrap_format)
        summary_sheet.set_row(0, 50)
        summary_sheet.set_column("A:A", 60)
        summary_sheet.set_column("B:B", 18, count_format)

        # remove Unnamed columns
        dfcols = list(sheet2.columns)
        for c in sheet2.columns:
            if isinstance(c, str) and "Unnamed" in c:
                dfcols.remove(c)
    
        sheet2 = sheet2.loc[:, dfcols]
        
        df_null = sheet2.isnull().sum()
        df_nnull = sheet2.notnull().sum()
        null_pct = list(range(0, len(df_null)))
        
    
        for i in range(0, len(df_null)):
            if df_null[i] + df_nnull[i] == 0:
                null_pct[i] = 0
            else:
                null_pct[i] = df_null[i] / (df_null[i] + df_nnull[i]) 
                if null_pct[i] > null_change/100:
                    di_6_3.collect_stats(aco_name, name1, dfcols[i] + f" % of Null > {null_change}%")
    
        freq = pandas.DataFrame({
              'Varaibles in File': dfcols,
              '# of Not Null':df_nnull,
              '# of Null': df_null,
              '% of Null': null_pct
                    })
       
        freq.to_excel(summary_writer, sheet_name = name1 + "_missing", index=False, startrow=1)
        freq_sheet = summary_writer.sheets[name1 + "_missing"]
        freq_sheet.merge_range('A1:D1', header, wrap_format)
        freq_sheet.set_row(0, 50)
        freq_sheet.set_column("A:A", 40, left_align_format)
        freq_sheet.set_column("B:B", 18, count_format)
        freq_sheet.set_column(2, 2, 18, count_format)
        freq_sheet.set_column(3, 3, 18, pct_format)
        
        if model == 'PROGRAM1':
            report_generated.set_header(header)
        
    diff_writer.save()
    summary_writer.save()
        
