# -*- coding: utf-8 -*-
"""
Created on Sat Oct 10 15:21:37 2020

@author: heart
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas


class DataInconsist_1_2:
    def __init__(self):
        # initialize data dictionary
        self.di = {
            'PPO': [], 
            'Validation Type': [],
            'Validation Error': []}
        self.header = None

    ### Read in the results from another program
    def collect_stats(self, freq, Excel_IDR_freq, exclusion_reasons_df, PPO, sk):
        for row in freq.index:
            if freq.loc[row, 'Compared Results'] == 'FALSE':
                self.di['Validation Type'].append('Summary_Of_Missing_Values')
                self.di['Validation Error'].append(freq.loc[row, 'Variables in File'])
                self.di['PPO'].append(PPO)
    
        for row in Excel_IDR_freq.index:
            if Excel_IDR_freq.loc[row, 'Compared Results'] == False:
                self.di['Validation Type'].append('Excel_vs_IDR_Summary')
                self.di['Validation Error'].append(Excel_IDR_freq.loc[row, 'Data Variables'])
                self.di['PPO'].append(PPO)
    
# =============================================================================
#         for row in df_excl.index:
#             if df_excl.loc[row, 'Excel VS IDR Results'] == 'Not Match':
#                 self.di['Validation Type'].append('Exclusions')
#                 self.di['Validation Error'].append(df_excl.loc[row, 'Bene Exclusion Reason Type'] + "_Excel/IDR")
#                 self.di['PPO'].append(PPO)
#             if df_excl.loc[row, 'Excel VS DB2 Results'] == 'Not Match':
#                 self.di['Validation Type'].append('Exclusions')
#                 self.di['Validation Error'].append(df_excl.loc[row, 'Bene Exclusion Reason Type'] + "_Excel/DB2")
#                 self.di['PPO'].append(PPO)
# =============================================================================

      
        for row in exclusion_reasons_df.index:
            if exclusion_reasons_df.loc[row, 'Compared Results'] == 1:
                self.di['Validation Type'].append('Exclusion_Reason_Types_Check')
                self.di['Validation Error'].append(exclusion_reasons_df.loc[row, 'Exclusion Reasons'])
                self.di['PPO'].append(PPO)

    ### Define header
    def set_header(self, header):
        if self.header is None:
            self.header = header

    ### Write the results into an Excel file
    def write(self, outfile):
        df = pandas.DataFrame(self.di)
        
        with pandas.ExcelWriter(outfile) as writer:
        
            df.to_excel(writer, sheet_name='1-2 Data Inconsistency Report', index=None, startrow=1)
    
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
    
            worksheet = writer.sheets['1-2 Data Inconsistency Report']
            worksheet.set_column(1, 0, 30)
            worksheet.set_column(2, 1, 60)
            worksheet.set_column(3, 2, 100)
            
            worksheet.merge_range('A1:D1', self.header, wrap_format)
            worksheet.set_row(0, 60)

