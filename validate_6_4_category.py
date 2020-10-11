# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 15:04:33 2020

@author: heart
"""


import pandas
import re


class Category_6_4:
    def __init__(self):
        self.aco_list = []
        self.names = []
        self.counts = []
        self.mis_v = [] 

        
    def validate_6_4_category(self, infile2):
        
        excel2 = pandas.read_excel(infile2, sheet_name=0, convert_float = False, header=5, index_col=None)
    
        #excel1.sort_values("MBI", inplace=True, )
        #excel2.sort_values("MBI", inplace=True)
    
    
        pattern = re.compile(r'.*(\d{4})(\d\d)(\d\d)\.xlsx')
        m = pattern.match(infile2)
        if m is None:
            print('File name does not match')
            quit()
            
        self.run_date = "Run Date:  " + m.group(2) + "/" + m.group(3) + "/" + m.group(1)
        
    ### Get report Number and header 
        excel0 = pandas.read_excel(infile2, 
                                   sheet_name=0, 
                                   nrows = 4,
                                   convert_float = True, 
                                   header=None, 
                                   index_col=None, 
                                   dtype="object")
    
    #    py0 = excel0.iloc[0,11]
    #    performance_year = int(py0[-4:])
        
     # Reporting Month: Apr, 2020		Reporting Month Start Date: 04/01/2020		Reporting Month End Date: 04/30/2020			Performance Year: 2020		
    
        self.report_name = excel0.iloc[0,0]
        report_aco = excel0.iloc[3,0]
        aco = report_aco[8:]
        
        category_list = ['A', 'D', 'E', 'N', 'V', 'X']
        cols_list = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']

    
        for col in range(8,len(excel2.columns)-1):
            self.aco_list.append(aco)
            name = excel2.columns[col]
            name = "Column " + cols_list[col - 8] + " - " + str(int(name)) 
            self.names.append(name)
            
            c = 0
            l = {}
            for row in excel2.index:
                v = str(excel2.iloc[row, col])
                if v not in category_list:
                    c += 1
                    if v == 'nan' or v == "":
                        v = "null"
                    l[v] = 1
                    
            self.counts.append(c)
            
            s = ""
            for k in l:
                if s == "":
                    s = k
                else:
                    s += ", " + k
                    
            self.mis_v.append(s)

     
    def save(self, outfile, sk):        
        csk = sk.get_csk()
        
        df = pandas.DataFrame({
                "PPO ID": self.aco_list,
                "Field validation on code values(column I to T)": self.names,
                "Mismatched Count": self.counts,
                "Mismatched Values": self.mis_v})    
            
    
    
        with pandas.ExcelWriter(outfile) as writer:
        
            # summary sheet
            df.to_excel(writer, sheet_name = "Summary of Categories", index=False, startrow=1)
        
            workbook = writer.book
            count_format = workbook.add_format({"num_format": "#,##0"})
            left_align_format = workbook.add_format()
            left_align_format.set_align("left")
            wrap_format = workbook.add_format()
            wrap_format.set_text_wrap()
        
            worksheet = writer.sheets["Summary of Categories"]
            worksheet.merge_range('A1:D1', self.report_name + "\n" + self.run_date + f"\nCurrent SK used: {csk}", wrap_format)
            worksheet.set_row(0, 50)
            worksheet.set_column("A:A", 30)
            worksheet.set_column("B:B", 60)
            worksheet.set_column("C:C", 30, count_format)
            worksheet.set_column("D:D", 30)
    
