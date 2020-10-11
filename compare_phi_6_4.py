# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 14:59:05 2020

@author: heart
"""


import pandas
import re

        
def compare_phi_6_4(infile1, infile2, outfile, sheet, header, sk):
    
    excel1 = pandas.read_excel(infile1, sheet_name=sheet, convert_float = False, header=header, index_col=None)
    excel2 = pandas.read_excel(infile2, sheet_name=sheet, convert_float = False, header=header, index_col=None)

    #excel1.sort_values("MBI", inplace=True, )
    #excel2.sort_values("MBI", inplace=True)

    csk = sk.get_csk()
    psk = sk.get_psk()
    
### Get report Number 
    excel0 = pandas.read_excel(infile2, 
                               sheet_name=sheet, 
                               nrows = 1,
                               convert_float = True, 
                               header=None, 
                               index_col=None, 
                               dtype="object")

#    py0 = excel0.iloc[0,11]
#    performance_year = int(py0[-4:])
    
 # Reporting Month: Apr, 2020		Reporting Month Start Date: 04/01/2020		Reporting Month End Date: 04/30/2020			Performance Year: 2020		

    report_name = excel0.iloc[0,0]
    
    pattern = re.compile(r'.*(\d{4})(\d\d)(\d\d)\.xlsx')
    m = pattern.match(infile2)
    if m is None:
        print('File name does not match')
        quit()
        
    run_date = "Run Date:  " + m.group(2) + "/" + m.group(3) + "/" + m.group(1)
        
    rows1 = excel1.index.size
    rows2 = excel2.index.size
    
    both_df = pandas.merge(excel1, excel2, on="Current MBI", how="inner")
    both_rows = both_df.index.size
    
    first_name_change = 0
    last_name_change = 0
    gender_change = 0
    
    for row in range(0, both_rows):
        if (both_df.at[row, "First Name_x"] != both_df.at[row, "First Name_y"]):
            first_name_change += 1
            
        if (both_df.at[row, "Last Name_x"] != both_df.at[row, "Last Name_y"]):
            last_name_change += 1
            
        if (both_df.at[row, "Gender_x"] != both_df.at[row, "Gender_y"]):
            gender_change += 1
            
    
    summary = pandas.DataFrame({
            "Type": ["Entitled Beneficiaries in Previous Month",
                     "Entitled Beneficiaries in Current Month",
                     "Entitled Beneficiaries in Both Months",
                     "Entitled Beneficiaries in Previous Month only",
                     "Entitled Beneficiaries in Current Month only",
                     "In Both Files: First Name Changed",
                     "In Both Files: Last Name Changed",
                     "In Both Files: Gender Changed"],
            "Count": [rows1, rows2, both_rows, rows1 - both_rows, rows2 - both_rows,
                      first_name_change, last_name_change, gender_change]
            })
    

    
    # remove Unnamed columns
    dfcols = list(excel2.columns)
    for c in excel2.columns:
        if isinstance(c, str) and "Unnamed" in c:
            dfcols.remove(c)

    excel2 = excel2.loc[:, dfcols]

    # remove duplicate records in current month
    excel2.drop_duplicates('Current MBI', inplace=True)

    excel2.dropna(how='all', subset=['Current MBI'], inplace=True)

    df_null = excel2.isnull().sum()
    df_nnull = excel2.notnull().sum()
    null_pct = list(range(0, len(df_null)))
    

    for i in range(0, len(df_null)):
        null_pct[i] = df_null[i] / (df_null[i] + df_nnull[i]) 

    freq = pandas.DataFrame({
          'Varaibles in File': dfcols,
          '# of Not Null':df_nnull,
          '# of Null': df_null,
          '% of Null': null_pct
                })

### Get the running date from file name infile2

    with pandas.ExcelWriter(outfile) as writer:
    
        # summary sheet
        summary.to_excel(writer, sheet_name = "Summary", index=False, startrow=1)
    
        workbook = writer.book
        count_format = workbook.add_format({"num_format": "#,##0"})
        left_align_format = workbook.add_format()
        left_align_format.set_align("left")
        wrap_format = workbook.add_format()
        wrap_format.set_text_wrap()
    
        worksheet = writer.sheets["Summary"]
        worksheet.merge_range('A1:B1', report_name + "\n" + run_date + f"\nCurrent SK used: {csk}", wrap_format)
        worksheet.set_row(0, 50)
        worksheet.set_column("A:A", 60)
        worksheet.set_column("B:B", 18, count_format)

        count_format = workbook.add_format({"num_format": "#,##0"})
        pct_format = workbook.add_format({"num_format": "0.00%"})
    
        freq.to_excel(writer, sheet_name = "Summary Of Missing Values", index=False, startrow=1)     
        summary_sheet = writer.sheets["Summary Of Missing Values"]   
        summary_sheet.merge_range('A1:D1', report_name + "\n" + run_date + f"\nCurrent SK used: {csk}" , wrap_format)
        summary_sheet.set_row(0, 50)
        summary_sheet.set_column("A:A", 40, left_align_format)
        summary_sheet.set_column("B:B", 18, count_format)
        summary_sheet.set_column(2, 2, 18, count_format)
        summary_sheet.set_column(3, 3, 18, pct_format)
