# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 14:54:42 2020

@author: heart
"""

import pandas



class DataInconsist_6_3:
    def __init__(self):
        # initialize data dictionary
        self.ppo = []
        self.validation_type = []
        self.validation_error = []
        self.header = None

### from Summary read in previous_records, current_records. If the change is greater than 2%(absolute value), then write to here for alert
### from freq, read in any % of Nulls > 3%. dfcols, null_pct

    def collect_stats(self, ppo, validation_type, validation_error):
        self.ppo.append(ppo)
        self.validation_type.append(validation_type)
        self.validation_error.append(validation_error)

    def set_header(self, header):
        if self.header is None:
            self.header = header


    def write(self, outfile):
        df = pandas.DataFrame({'PPO': self.ppo,
                               'Validation Type': self.validation_type,
                               'Validation Error': self.validation_error})
        
        with pandas.ExcelWriter(outfile) as writer:
        
            df.to_excel(writer, sheet_name='6-3 Data Inconsistency Report', index=None, startrow=1)
    
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
    
            worksheet = writer.sheets['6-3 Data Inconsistency Report']
            worksheet.set_column(1, 0, 30)
            worksheet.set_column(2, 1, 60)
            worksheet.set_column(3, 2, 100)
            
            worksheet.merge_range('A1:D1', self.header, wrap_format)
            worksheet.set_row(0, 60)

