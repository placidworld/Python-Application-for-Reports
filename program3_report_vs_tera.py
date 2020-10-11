# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 15:10:53 2020

@author: heart
"""


import pandas
import sqlalchemy
import re

from PRACTICE.sys_connect import hostname, username, password, database, IDR_host, path, outpath

pandas.set_option('display.float_format', lambda x: '%.3f' % x)


def program3_report_vs_tera(infile2, outfile):
    
    print("Read in SSP current month data ...")

    #### Read in SSP reports from IDR
    
    sql = f"""
            """

    engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
    excel1 = pandas.read_sql(sql,con=engine)
      

    
    excel2 = pandas.read_excel(infile2, convert_float = True, header=None, index_col=None, dtype="object", skiprows=[0, 1, 2, 3, 4])

    excel0 = pandas.read_excel(infile2, 
                               sheet_name=0, 
                               convert_float = True, 
                               header=None, 
                               index_col=None, 
                               dtype="object")

    report_name = excel0.iloc[1,0]
    report_period = excel0.iloc[2,0]
    report_date = excel0.iloc[-1,0]
    
    
#    excel1.drop(columns=[10, 11, 16, 17], inplace=True)
    excel2.drop(columns=[10, 11, 16, 17], inplace=True)
    
    excel1.rename(columns={
            "PERFORMANCE_YEAR": "Performance Year",
            "BENE_CNT": "Beneficiaries",
            "TOT_SA_CLM_PCT": "% of Total SA Expenditures",
            "TOT_SA_CLM_CNT": "SA Claims Suppressed",
            "SPRSD_SA_PCT": "% of SA Claims Suppressed",
            "PBP_RDCTN_BENE_CNT": "Beneficiaries with PBP Reduction",
            "PBP_INCLSN_TOT_AMT": "Total PBP Inclusion Amount",
            "PBP_INCLSN_RDCTN_PCT": "% of Total PBP Reduction Amount",
            "TOT_UNIQUE_BENE": "Total Count of Unique Beneficiaries"
            }, inplace=True)
    
    excel2.rename(columns={
            0: "PPO",
            1: "Performance Year",
            2: "Beneficiaries",
            3: "% of Total SA Expenditures",
            4: "SA Claims Suppressed",
            5: "% of SA Claims Suppressed",
            6: "Beneficiaries with PBP Reduction",
            7: "Total PBP Inclusion Amount",
            8: "% of Total PBP Reduction Amount",
            9: "BD",
            12: "PL",
            13: "EC",
            14: "PC",
            15: "Total Count of Unique Beneficiaries"
            }, inplace = True)
    
    ppo_pat = re.compile('^A[0-9]+$')
    
    for row in excel1.index:
        excel1.iloc[row, 0] = str(excel1.iloc[row, 0]).strip()
        excel1.iloc[row, 3] = round(excel1.iloc[row, 3] * 100)
        excel1.iloc[row, 5] = round(excel1.iloc[row, 5] * 100)
        excel1.iloc[row, 8] = round(excel1.iloc[row, 8] * 100)
        if not ppo_pat.match(str(excel1.iloc[row, 0])):
            excel1.drop(range(row, len(excel1.index)), inplace = True)
            break

    excel1.sort_values("ACO", inplace=True)
        
    for row in excel2.index:
        excel2.iloc[row, 3] = round(excel2.iloc[row, 3] * 10000)
        excel2.iloc[row, 5] = round(excel2.iloc[row, 5] * 10000)
        excel2.iloc[row, 8] = round(excel2.iloc[row, 8] * 10000)
        if not ppo_pat.match(str(excel2.iloc[row, 0])):
            excel2.drop(range(row, len(excel2.index)), inplace = True)
            break

    excel2.sort_values("ACO", inplace=True)
            
    both_df = pandas.merge(excel1, excel2, on="ACO", how="inner")
    
    num_pattern = re.compile("^[-0-9.]+$")
    currency_pattern = re.compile("[-0-9,$]+\.[0-9][0-9]?$")
    
    with pandas.ExcelWriter(outfile) as writer:
        work_book = writer.book
        left_align_format = work_book.add_format()
        left_align_format.set_align("left")
        red_format = work_book.add_format()
        red_format.set_font_color('red')
        count_format = work_book.add_format({"num_format": "#,##0"})
        pct_format = work_book.add_format({"num_format": "0.00%"})
        currency_format = work_book.add_format({"num_format": "$#,##0.00"})
        wrap_format = work_book.add_format()
        wrap_format.set_text_wrap()
               
        diff_sheet = work_book.add_worksheet("Compared result")
        diff_sheet.merge_range('A1:N1', report_name + "\n" + report_period + "\n" + report_date, wrap_format)
        diff_sheet.set_row(0, 45)

        rows = len(both_df.index)
        cols = len(excel1.columns)
        
        # print header
        for col in range(0, cols):
            diff_sheet.write(1, col, excel1.columns[col])
        
        for row in range(0, rows):
            
            # write ACO and Performance year
            diff_sheet.write(row + 2, 0, str(both_df.iloc[row, 0]))
            diff_sheet.write(row + 2, 1, "{:.0f}".format(both_df.iloc[row, 1]))
            
            for col in range(2, cols):
                a = both_df.iloc[row, col]
                b = both_df.iloc[row, col + cols - 1]

                if a == b:
                    diff_sheet.write(row + 2, col, "Matched")
                else:
                    diff_sheet.write(row + 2, col, "Not Matched", red_format)
 
        diff_sheet.set_column(0, cols, 30, left_align_format)

