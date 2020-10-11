# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 14:06:04 2020

@author: heart
"""


import pandas
import ibm_db


from PRACTICE.sys_connect import hostname, username, password, database, IDR_host

###
class Compare_6_1():
    def __init__(self):
        self.aco_list = []
        self.category_list = []
        self.report_list = []
        self.db2_list = []
        self.results_list = []
        self.rec_updt_dt = '01/01/1970'
        

    def compare_6_1(self, infile1, infile2, model, aco):
        print("Comparing 6-1 report now ...")
    
        excel0 = pandas.read_excel(infile2, 
                                   sheet_name=0, 
                                   convert_float = True, 
                                   header=None, 
                                   index_col=None, 
                                   skiprows=[0, 1, 2],
                                   dtype="object")
    
        excel_aligned_bene_count = excel0.iloc[0, 3]
        
    #    PROGRAM261_BENE_COUNT                      
        excel_partA_claims_amt = excel0.iloc[18, 16]
        excel_partB_claims_amt = excel0.iloc[24, 16]
        
        if model == 'PROGRAM2':    
            sql = f"""
                   
                     WITH UR;
               """ 
        else:
            sql = f"""
               """ 
            
        
    #    audit_aligned_bene_count = []
    #    audit_partA_claims_amt = []
     #   audit_partB_claims_amt = []
     
        audit_data = []
    
   
        conn = ibm_db.pconnect(f"DATABASE={database};HOSTNAME={hostname};PORT=12003;PROTOCOL=TCPIP;UID={username};PWD={password}",'','')
        stmt = ibm_db.prepare(conn,sql)
        
        print("Running SQL query ...")
        ibm_db.execute(stmt)
        print("Running SQL query ... done")
        
        result_dict = ibm_db.fetch_assoc(stmt)
        
        while result_dict:
                audit_data.append(result_dict)
                result_dict = ibm_db.fetch_assoc(stmt)
            
        audit_aligned_bene_count = 0
        audit_partA_claims_amt = 0        
        audit_partB_claims_amt = 0
        
        audit_df = pandas.DataFrame(audit_data)
        
        for r in audit_df.index:
            column_name = audit_df.loc[r, 'CLMN_NAME'].strip()
            stats_cnt = audit_df.loc[r, 'STATS_CNT']
            self.rec_updt_dt = audit_df.loc[r, 'REC_UPDT_DT']
            if column_name == 'PROGRAM261_BENE_COUNT' or column_name == 'PROGRAM161_BENE_COUNT':
                audit_aligned_bene_count = stats_cnt
            elif column_name == 'PROGRAM261_PARTA_AMT' or column_name == 'PROGRAM161_PARTA_AMT':
                audit_partA_claims_amt = stats_cnt
            elif column_name == 'PROGRAM261_PARTB_AMT' or column_name == 'PROGRAM161_PARTB_AMT':
                audit_partB_claims_amt = stats_cnt
        
        if int(excel_aligned_bene_count) == audit_aligned_bene_count or int(excel_aligned_bene_count) < audit_aligned_bene_count:
            aligned_bene_count_match = 'MATCH'
        else:
            aligned_bene_count_match = 'NOT MATCH'  
    
        if int(excel_partA_claims_amt) == audit_partA_claims_amt or int(excel_partA_claims_amt) < audit_partA_claims_amt:
            partA_claims_amt_match = 'MATCH'
        else:
            partA_claims_amt_match = 'NOT MATCH'  
    
        if int(excel_partB_claims_amt) == audit_partB_claims_amt or int(excel_partB_claims_amt) > audit_partB_claims_amt:
            partB_claims_amt_match = 'MATCH'
        else:
            partB_claims_amt_match = 'NOT MATCH' 
            
        self.aco_list.extend([aco, aco, aco])
        self.category_list.extend(["Total Count of Aligned Beneficiaries", 'Total Part A Claims Amount', 'Total Part B Claims Amount'])
        self.report_list.extend([excel_aligned_bene_count, excel_partA_claims_amt, excel_partB_claims_amt])
        self.db2_list.extend([audit_aligned_bene_count, audit_partA_claims_amt, audit_partB_claims_amt])
        self.results_list.extend([aligned_bene_count_match, partA_claims_amt_match, partB_claims_amt_match])
        
        
            
    def compare_6_1_save(self, audit_file):
    
        rpt = pandas.DataFrame({'PPO ID': self.aco_list,
                                'Category' : self.category_list,
                                'Report': self.report_list,
                                'DB2 Audit Table': self.db2_list,
                               'Compared Results': self.results_list})    
    
    
        audit_run_dt = 'Audit Table Updated Date:  ' + self.rec_updt_dt
        writer = pandas.ExcelWriter(audit_file)
    
        rpt.to_excel(writer, sheet_name='6-1 Audit Validation', index=None, startrow=1)
            
        work_book = writer.book
        left_align_format = work_book.add_format()
        left_align_format.set_align("left")
        right_align_format = work_book.add_format()
        right_align_format.set_align("right")
        count_format = work_book.add_format({"num_format": "#,##0"})
        red_format = work_book.add_format()
        red_format.set_font_color('red')
        bold_format = work_book.add_format({'bold': True})
        merge_format = work_book.add_format({'align': 'left'})
        wrap_format = work_book.add_format()
        wrap_format.set_text_wrap()
    
        full_border = work_book.add_format(
                      {"border": 1,
                      "border_color": "#000000"
                      })
        
        for worksheet in writer.sheets.values():                
            worksheet.merge_range('A1:E1', audit_run_dt, wrap_format)
            worksheet.set_row(0, 20)
            worksheet.set_column("A:A", 20)
            worksheet.set_column("B:B", 48)
            worksheet.set_column("C:C", 30, count_format)
            worksheet.set_column(3, 3, 30, count_format)
            worksheet.set_column(4, 4, 30, right_align_format)
    
        writer.save()
    
