# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 15:16:51 2020

@author: heart
"""

# Read in monthly SK # based on verbal/written notification

# N12, V12, N51A, N51B, V51A, V51B, N61, V61

import os
import pandas
import numpy

import re
import ibm_db

from PRACTICE.sys_connect import hostname, username, password, database, IDR_host, path, outpath

pandas.set_option('display.float_format', lambda x: '%.3f' % x)


    
class PROGRAM1_6_3_Report_Generated():

    def __init__(self):
        self.file_ppo_dict = {}
        self.sql_ppo_list = []
        
    def collect_ppo_names(self, model, report_num, current_month):

        pattern = re.compile("^" + model + "_" + report_num + "_ ?([A-Z]\d\d\d)_" + current_month + "\d\d\.")
        
        files = os.listdir(path)
        files.sort()
        for fname in files:
                        
            m = pattern.match(fname)
            if not m:
                continue
            
            print("matched:" + fname)
                            
            ppo = m.group(1)
            self.file_ppo_dict[ppo] = 1
        
    def query_ppo(self):
        # Retrieve Program1 ppos with PBP/AIPBP from db2
        sql = f"""
            """      
        
        db2_data = []

        conn = ibm_db.pconnect(f"DATABASE={database};HOSTNAME={hostname};PORT=12003;PROTOCOL=TCPIP;UID={username};PWD={password}",'','')
        stmt = ibm_db.prepare(conn,sql)
        
        print("Running SQL query ...")
        ibm_db.execute(stmt)
        print("Running SQL query ... done")
        
        result_dict = ibm_db.fetch_assoc(stmt)
        
        while result_dict:
            db2_data.append(result_dict)
            result_dict = ibm_db.fetch_assoc(stmt)
    
            
        db2_df = pandas.DataFrame(db2_data)
        self.sql_ppo_list = db2_df.loc[:, 'ppo']
    
    def set_header(self, header):
        self.header = header
        
        
    def gen_report(self, outfile):
        file_ppo_list = []
        result_list = []
        
        for ppo in self.sql_ppo_list:
            ppo = ppo.strip()
            if ppo in self.file_ppo_dict:
                result_list.append('Generated')
                file_ppo_list.append(ppo)
            else:
                result_list.append('Not Generated')
                file_ppo_list.append("")
            
        df = pandas.DataFrame({
                'SQL Result': self.sql_ppo_list,
                'Report Generated': file_ppo_list,
                'Outcome': result_list
                })
    
        with pandas.ExcelWriter(outfile) as writer:
            df.to_excel(writer, sheet_name = '6_3_ppos_Report_Generated', index=False, startrow = 1)
            
            workbook = writer.book
            right_format = workbook.add_format({'align': 'right'})
            wrap_format = workbook.add_format()
            wrap_format.set_text_wrap()
            
            df_sheet = writer.sheets["6_3_ppos_Report_Generated"]
            df_sheet.merge_range('A1:C1', self.header, wrap_format)
            df_sheet.set_row(0, 50)
            df_sheet.set_column("A:A", 30)
            df_sheet.set_column(1, 1, 30, right_format)
            df_sheet.set_column(2, 2, 30, right_format)
        

