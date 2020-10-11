# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 15:20:17 2020

@author: heart
"""


# Read in monthly SK # based on verbal/written notification

# N12, V12, N51A, N51B, V51A, V51B, N61, V61

import pandas
import sqlalchemy


#from IDDOC.sys_connect import path

from PRACTICE.sys_connect import hostname, username, password, database, IDR_host, path


    
class PROGRAM1_PROGRAM2_SK:
    def __init__(self):
       pass
   
    def query(self, model):

        # performance_year = 2020
        # For PROGRAM2 1-2, 5-1, 6-1, 6-4

        if model == 'PROGRAM2':
            sql = f"""
                    """
        else:
            sql = f"""
                  
            """      
            
    
        engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
        csk_data = pandas.read_sql(sql,con=engine)
         
    
        #### Get previous month SK 
                
        if model == 'PROGRAM2':
            sql = f"""
                  
                   """
        else:
            sql = f"""
                  
            """      
                
        engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
        psk_data = pandas.read_sql(sql,con=engine)
        
        csk0 = csk_data.at[0,'csk']
        psk0 = psk_data.at[0, 'psk']
        
        csk = input(f"Enter Current SK [{csk0}]:  ")
        if csk.strip() == "":
            self.csk = csk0
        else:
            self.csk = int(csk)
            
        psk = input(f"Enter Previous SK [{psk0}]:  ")
        if psk.strip() == "":
            self.psk = psk0
        else:
            self.psk = int(psk)

    def get_csk(self):
        return self.csk

    def get_psk(self):
        return self.psk        

