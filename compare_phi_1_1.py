# -*- coding: utf-8 -*-
"""
Created on Sat Oct 10 10:35:14 2020

@author: heart
"""

import ibm_db
import sqlalchemy
import numpy

from PRACTICE.sys_connect import hostname, username, password, database, IDR_host, path

import pandas
import re

### Check entire database and convert all Empty values to Null
def convert_empty_to_null(df):
    blank_pat = re.compile("^[ \t\r\n]*$")    
    for row in df.index:
        for col in df.columns:
            v = df.loc[row, col]
            if isinstance(v, str) and blank_pat.match(v):
                df.loc[row, col] = numpy.NaN
            elif v is None:
                df.loc[row, col] = numpy.NaN

def compare_phi(infile1, infile2, summary_file, diff_file, diff_file1, model, aco_name, sheet, header, sk, di_1_2, error_threshold):
    
    print("Calculating phi ...")
    
    ### Read in two month/quarter/year Excel files 
    ### Excel1 is previous month/quarter/year
    ### Excel2 is current mont/quarter/year
    ### pandas.read_excel to create dataframe from excel files 
    excel1 = pandas.read_excel(infile1, sheet_name=sheet, convert_float = False, header=header, index_col=None)
    excel2 = pandas.read_excel(infile2, sheet_name=sheet, convert_float = False, header=header, index_col=None)

    ### Read in excel2 header information for the validation report to use
    excel0 = pandas.read_excel(infile2, 
                               sheet_name=sheet, 
                               nrows = 3,
                               convert_float = True, 
                               header=None, 
                               index_col=None, 
                               dtype="object")

#    py0 = excel0.iloc[0,11]
#    performance_year = int(py0[-4:])
    
 # Reporting Month: Apr, 2020		Reporting Month Start Date: 04/01/2020		Reporting Month End Date: 04/30/2020			Performance Year: 2020		
    ### Read in different values for output file header 
    report_gen = excel0.iloc[1,1]
    report_desc = excel0.iloc[2,1]
    report_gen = report_gen.replace("\n ", "\n")
    report_gen = report_gen.replace("\n ", "\n")
    report_gen = report_gen.replace("\n ", "\n")
    
#### Rename Excel Merged Columns for later validation    
    excel1.rename(columns = {
            'Address': 'Address_Line1',
            'Unnamed: 7': 'Address_Line2',
            'Unnamed: 8': 'Address_Line3',
            'Unnamed: 9': 'Address_Line4',
            'Unnamed: 10': 'Address_Line5',
            'Unnamed: 11': 'Address_Line6',
            }, inplace = True)
 
    excel2.rename(columns = {
            'Address': 'Address_Line1',
            'Unnamed: 7': 'Address_Line2',
            'Unnamed: 8': 'Address_Line3',
            'Unnamed: 9': 'Address_Line4',
            'Unnamed: 10': 'Address_Line5',
            'Unnamed: 11': 'Address_Line6',
            }, inplace = True)
    
    #excel1.sort_values("MBI", inplace=True, )
    #excel2.sort_values("MBI", inplace=True)

    # remove Unnamed columns from both Excel1 and Excel2 - orginal reports has a lot merged lines/columns
    ### list(excel2.columns) create a list consists of all column names
    ### Then use for loop to check if any columns name contains "Unnamed"
    ### use isinstance to check 
    ### if there is a column name has "Unnamed" in it, then use remove function to remove it
    dfcols = list(excel2.columns)
    for c in excel2.columns:
        if isinstance(c, str) and "Unnamed" in c:
            dfcols.remove(c)

    ### replace columns with the newly generated Column list
    excel2 = excel2.loc[:, dfcols]

    # remove duplicate records in current month
    # remove missing values, use option how = 'all', means any empty line will be remvoed
    excel1.drop_duplicates('MBI', inplace=True)
    excel1.dropna(how='all', subset=['MBI'], inplace=True)
    excel2.drop_duplicates('MBI', inplace=True)
    excel2.dropna(how='all', subset=['MBI'], inplace=True)
    #convert_empty_to_null(excel2)
    
    ### Count the total records 
    rows1 = excel1.index.size
    rows2 = excel2.index.size
    
    ### Merge the files with inner merge
    both_df = pandas.merge(excel1, excel2, on="MBI", how="inner")
    both_rows = both_df.index.size
    
    first_name_change = 0
    last_name_change = 0
    addr_change = 0
    gender_change = 0
    date_change = 0
    reason_change = 0
    
    for row in range(0, both_rows):
        if (str(both_df.at[row, "First Name_x"]) != str(both_df.at[row, "First Name_y"])):
            first_name_change += 1
            
        if (str(both_df.at[row, "Last Name_x"]) != str(both_df.at[row, "Last Name_y"])):
            last_name_change += 1
            
        if (str(both_df.at[row, "Address_Line1_x"]) != str(both_df.at[row, "Address_Line1_y"]) or 
            str(both_df.at[row, "Address_Line2_x"]) != str(both_df.at[row, "Address_Line2_y"]) or
            str(both_df.at[row, "Address_Line3_x"]) != str(both_df.at[row, "Address_Line3_y"]) or
            str(both_df.at[row, "Address_Line4_x"]) != str(both_df.at[row, "Address_Line4_y"]) or
            str(both_df.at[row, "Address_Line5_x"]) != str(both_df.at[row, "Address_Line5_y"]) or
            str(both_df.at[row, "Address_Line6_x"]) != str(both_df.at[row, "Address_Line6_y"]) or
            str(both_df.at[row, "State_x"]) != str(both_df.at[row, "State_y"]) or
            str(both_df.at[row, "Zip Code_x"]) != str(both_df.at[row, "Zip Code_y"])):
            addr_change += 1
            
        if (str(both_df.at[row, "Gender_x"]) != str(both_df.at[row, "Gender_y"])):
            gender_change += 1
            
        if (str(both_df.at[row, "Date of Exclusion_x"]) != str(both_df.at[row, "Date of Exclusion_y"])):
            date_change += 1

        if (str(both_df.at[row, "Reason for Exclusion_x"]) != str(both_df.at[row, "Reason for Exclusion_y"])):
            reason_change += 1
            
        ### Create summary dataframe for report use
        summary = pandas.DataFrame({
             "Type": ["Excluded Beneficiaries in Previous Month",
                      "Excluded Beneficiaries in Current Month",
                      "Excluded Beneficiaries in Both Months",
                      "Excluded Beneficiaries in Previous Month only",
                      "Excluded Beneficiaries in Current Month only",
                      "In Both Files: First Name Changed",
                      "In Both Files: Last Name Changed",
                      "In Both Files: Address Changed",
                      "In Both Files: Gender Changed",
                      "In Both Files: Date of Exclusion Changed",
                      "In Both Files: Reason for Exclusion Changed"],
             "Count": [rows1, rows2, both_rows, rows1 - both_rows, rows2 - both_rows,
                       first_name_change, last_name_change, addr_change, gender_change, date_change, reason_change]
                })

    keys = both_df["MBI"]
    only1 = excel1[~excel1["MBI"].isin(keys)]
    only2 = excel2[~excel2["MBI"].isin(keys)]
 

    df_null = excel2.isnull().sum()
    df_nnull = excel2.notnull().sum()
    null_pct = list(range(0, len(df_null)))
    

    for i in range(0, len(df_null)):
        null_pct[i] = df_null[i] / (df_null[i] + df_nnull[i]) 



#####################################################################
######## Retrieve the BENE level exclusions data from IDR 
######## Make sure to update the SK with the most recent one 
#####################################################################

    csk = sk.get_csk()
    psk = sk.get_psk()

# =============================================================================
#     if model == "PROGRAM1":
#         csk = sk_dic["CN12"]
#         psk = sk_dic["PN12"]
#     else:
#         csk = sk_dic["CV12"]
#         psk = sk_dic["PV12"]
#     
# =============================================================================
#### Current month IDR data
    if model == 'PROGRAM2':
        sql = f"""
               SELECT	*
                FROM ****                """
    ### initialize a list
    idr_excl_data = []

    ### Connecnt to Teradata    
    engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
    idr_excl_data = pandas.read_sql(sql,con=engine)
     
    idr_excl_df = pandas.DataFrame(idr_excl_data)
    
    #convert_empty_to_null(excel2)
    convert_empty_to_null(idr_excl_df)
    
    # strip tailing spaces in MBI
    for row in idr_excl_df.index:
        idr_excl_df.loc[row, 'MBI'] = str.strip(idr_excl_df.loc[row, 'MBI'])
        
    #### Compare Excel2 against IDR table 
    rows3 = idr_excl_df.index.size
    
    both_df_23 = pandas.merge(excel2, idr_excl_df, on="MBI", how="inner")
    both_rows_23 = both_df_23.index.size
    
#    MBI_change = 0
    HICNO_change = 0
    first_name_change = 0
    last_name_change = 0
    Address_Line1_change = 0
    Address_Line2_change = 0
    Address_Line3_change = 0
    Address_Line4_change = 0
    Address_Line5_change = 0
    Address_Line6_change = 0
    State_change = 0
    Zip_Code_change = 0
    Gender_change = 0
    Birth_Date_change = 0
    Date_of_Exclusion_change = 0
    Reason_for_Exclusion_change = 0
    
    change_dic = {'MBI': [],
                  'Data Source': [],
                  'HICNO': [],
                  'First Name': [],
                  'Last Name': [],
                  'Address Line 1': [],
                  'Address Line 2': [],
                  'Address Line 3': [],
                  'Address Line 4': [],
                  'Address Line 5': [],
                  'Address Line 6': [],
                  'State': [],
                  'Zip Code': [],
                  'Gender': [],
                  'Birth Date': [],
                  'Date of Exclusion': [],
                  'Reason For Exclusion': []
                  }
    
    for row in range(0, both_rows_23):
        
        changed = 0
        
        v1 = str(both_df_23.at[row, "HICNO_x"]).strip()
        v2 = str(both_df_23.at[row, "HICNO_y"]).strip()
        if (v1 != v2):
            HICNO_change += 1
            changed = 1
            
        v1 = str(both_df_23.at[row, "First Name"]).strip()
        v2 = str(both_df_23.at[row, "FIRST_NAME"]).strip()
        if (v1 != v2):
            first_name_change += 1
            changed = 1
            print("First name changed: MBI={}, excel={}, idr={}".format(both_df_23.loc[row, 'MBI'], v1, v2))
        
        v1 = str(both_df_23.at[row, "Last Name"]).strip()
        v2 = str(both_df_23.at[row, "LAST_NAME"]).strip()
        if (v1 != v2):
            last_name_change += 1
            changed = 1
            print("Last name changed: MBI={}, excel={}, idr={}".format(both_df_23.loc[row, 'MBI'], v1, v2))
            
        v1 = str(both_df_23.at[row, "Address_Line1_x"]).strip()
        v2 = str(both_df_23.at[row, "Address_Line1_y"]).strip()
        if (v1 != v2):
            Address_Line1_change += 1
            changed = 1
            print("Address line 1 changed: MBI={}, excel={}, idr={}".format(both_df_23.loc[row, 'MBI'], v1, v2))

        v1 = str(both_df_23.at[row, "Address_Line2_x"]).strip()
        v2 = str(both_df_23.at[row, "Address_Line2_y"]).strip()
        if (v1 != v2):
            Address_Line2_change += 1
            changed = 1
            print(f"Address line 2 changed: MBI={both_df_23.loc[row, 'MBI']}, excel={v1}, idr={v2}")
            
        v1 = str(both_df_23.at[row, "Address_Line3_x"]).strip()
        v2 = str(both_df_23.at[row, "Address_Line3_y"]).strip()
        if (v1 != v2):
            Address_Line3_change += 1 
            changed = 1

        v1 = str(both_df_23.at[row, "Address_Line4_x"]).strip()
        v2 = str(both_df_23.at[row, "Address_Line4_y"]).strip()
        if (v1 != v2):
            Address_Line4_change += 1 
            changed = 1
            print("Address line 4 changed: MBI={}, excel={}, idr={}".format(both_df_23.loc[row, 'MBI'], v1, v2))

        v1 = str(both_df_23.at[row, "Address_Line5_x"]).strip()
        v2 = str(both_df_23.at[row, "Address_Line5_y"]).strip()
        if (v1 != v2):
            Address_Line5_change += 1 
            changed = 1

        v1 = str(both_df_23.at[row, "Address_Line6_x"]).strip()
        v2 = str(both_df_23.at[row, "Address_Line6_y"]).strip()
        if (v1 != v2):
            Address_Line6_change += 1 
            changed = 1

        v1 = str(both_df_23.at[row, "State_x"]).strip()
        v2 = str(both_df_23.at[row, "State_y"]).strip()
        if (v1 != v2):
            State_change += 1
            changed = 1
            print("State value changed: MBI={}, excel={}, idr={}".format(both_df_23.loc[row, 'MBI'], v1, v2))
            if both_df_23.loc[row, 'MBI'] == '2HK2DN9EM96':
                pass

        v1 = str(both_df_23.at[row, "Zip Code"]).strip()
        v2 = str(both_df_23.at[row, "Zip_Code"]).strip()
        if (v1 != v2):
            Zip_Code_change += 1
            changed = 1
            print("Zip Code value changed: MBI={}, excel={}, idr={}".format(both_df_23.loc[row,'MBI'], v1, v2))
            
        v1 = str(both_df_23.at[row, "Gender_x"]).strip()
        v2 = str(both_df_23.at[row, "Gender_y"]).strip()
        if (v1 != v2):
            gender_change += 1
            changed = 1
            
        v1 = str(both_df_23.at[row, "Birth Date"].to_pydatetime().date())
        v2 = str(both_df_23.at[row, "Birth_Date"])
        if (v1 != v2):
            Birth_Date_change += 1
            changed = 1

        v1 = str(both_df_23.at[row, "Date of Exclusion"].to_pydatetime().date())
        v2 = str(both_df_23.at[row, "Exclusion_Date"])
        if (v1 != v2):
            Date_of_Exclusion_change += 1
            changed = 1

        v1 = str(both_df_23.at[row, "Reason for Exclusion"]).strip()
        v2 = str(both_df_23.at[row, "Exclusion_Reason"]).strip()
        if (v1 != v2):
            Reason_for_Exclusion_change += 1
            changed = 1
        
        if changed == 1:
            change_dic['MBI'].append(both_df_23.at[row, 'MBI'])
            change_dic['Data Source'].append('Excel')
            change_dic['HICNO'].append(both_df_23.at[row, "HICNO_x"])
            change_dic['First Name'].append(both_df_23.at[row, "First Name"])
            change_dic['Last Name'].append(both_df_23.at[row, "Last Name"])
            change_dic['Address Line 1'].append(both_df_23.at[row, "Address_Line1_x"])
            change_dic['Address Line 2'].append(both_df_23.at[row, "Address_Line2_x"])
            change_dic['Address Line 3'].append(both_df_23.at[row, "Address_Line3_x"])
            change_dic['Address Line 4'].append(both_df_23.at[row, "Address_Line4_x"])
            change_dic['Address Line 5'].append(both_df_23.at[row, "Address_Line5_x"])
            change_dic['Address Line 6'].append(both_df_23.at[row, "Address_Line6_x"])
            change_dic['State'].append(both_df_23.at[row, "State_x"])
            change_dic['Zip Code'].append(both_df_23.at[row, "Zip Code"])
            change_dic['Gender'].append(both_df_23.at[row, "Gender_x"])
            change_dic['Birth Date'].append(both_df_23.at[row, "Birth Date"])
            change_dic['Date of Exclusion'].append(both_df_23.at[row, "Date of Exclusion"])
            change_dic['Reason For Exclusion'].append(both_df_23.at[row, "Reason for Exclusion"])
            
            change_dic['MBI'].append(both_df_23.at[row, 'MBI'])
            change_dic['Data Source'].append('IDR')
            change_dic['HICNO'].append(both_df_23.at[row, "HICNO_y"])
            change_dic['First Name'].append(both_df_23.at[row, "FIRST_NAME"])
            change_dic['Last Name'].append(both_df_23.at[row, "LAST_NAME"])
            change_dic['Address Line 1'].append(both_df_23.at[row, "Address_Line1_y"])
            change_dic['Address Line 2'].append(both_df_23.at[row, "Address_Line2_y"])
            change_dic['Address Line 3'].append(both_df_23.at[row, "Address_Line3_y"])
            change_dic['Address Line 4'].append(both_df_23.at[row, "Address_Line4_y"])
            change_dic['Address Line 5'].append(both_df_23.at[row, "Address_Line5_y"])
            change_dic['Address Line 6'].append(both_df_23.at[row, "Address_Line6_y"])
            change_dic['State'].append(both_df_23.at[row, "State_y"])
            change_dic['Zip Code'].append(both_df_23.at[row, "Zip_Code"])
            change_dic['Gender'].append(both_df_23.at[row, "Gender_y"])
            change_dic['Birth Date'].append(both_df_23.at[row, "Birth_Date"])
            change_dic['Date of Exclusion'].append(both_df_23.at[row, "Exclusion_Date"])
            change_dic['Reason For Exclusion'].append(both_df_23.at[row, "Exclusion_Reason"])
    
    change_df = pandas.DataFrame(change_dic)
    
    with pandas.ExcelWriter(diff_file1) as change_writer:
        change_df.to_excel(change_writer, sheet_name='Changed Data Details', index=None)

    del change_df
    del change_dic
    
    

########################################################
#### Read in IDR previous month/qarter data 
#########################################################

    if model == 'PROGRAM2':
        sql = f"""
               """

    idr_excl_data1 = []
    
    try:
        engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
        idr_excl_data1 = pandas.read_sql(sql,con=engine)
     
    except Exception as e:
        print("SCRIPT: We have a problem")
        
    idr_excl_df1 = pandas.DataFrame(idr_excl_data1)
    convert_empty_to_null(idr_excl_df1)
    
    # strip tailing spaces in MBI
    for row in idr_excl_df1.index:
        idr_excl_df1.loc[row, 'MBI'] = str.strip(idr_excl_df1.loc[row, 'MBI'])
        
    #### Compare Excel2 against IDR table 
 #   rows4 = idr_excl_df1.index.size


    l = [HICNO_change / rows3, first_name_change / rows3, last_name_change / rows3, 
                                Address_Line1_change / rows3, Address_Line2_change / rows3,
                                Address_Line3_change / rows3, Address_Line4_change / rows3,
                                Address_Line5_change / rows3, Address_Line6_change / rows3,
                                State_change / rows3, Zip_Code_change / rows3,
                                Gender_change / rows3, Birth_Date_change / rows3,
                                Date_of_Exclusion_change / rows3, Reason_for_Exclusion_change / rows3]
    
    for i in range(0, len(l)):
        if l[i] * 100 > error_threshold:
            l[i] = "FALSE"
        else:
            l[i] = "TRUE"
              
### Count the # of Matched and Non Matched and % of Non-Matched 
    Excel_IDR_freq = pandas.DataFrame({
            'Data Variables': ['HICNO', 	'First Name', 'Last Name', 
                               'Address_Line1', 'Address_Line2',
                               'Address_Line3', 'Address_Line4', 
                               'Address_Line5', 'Address_Line6',
                               'State', 'Zip Code',
                               'Gender', 'Birth Date',
                               'Date of Exclusion', 'Reason for Exclusion'],
            '# of Not Matched': [HICNO_change, first_name_change, last_name_change, 
                                Address_Line1_change, Address_Line2_change,
                                Address_Line3_change, Address_Line4_change,
                                Address_Line5_change, Address_Line6_change,
                                State_change, Zip_Code_change,
                                Gender_change, Birth_Date_change,
                                Date_of_Exclusion_change, Reason_for_Exclusion_change
                                ],
            '# of Matched': [rows3 - HICNO_change, rows3 - first_name_change, rows3 - last_name_change, 
                                rows3 - Address_Line1_change, rows3 - Address_Line2_change,
                                rows3 - Address_Line3_change, rows3 - Address_Line4_change,
                                rows3 - Address_Line5_change, rows3 - Address_Line6_change,
                                rows3 - State_change, rows3 - Zip_Code_change,
                                rows3 - Gender_change, rows3 - Birth_Date_change,
                                rows3 - Date_of_Exclusion_change, rows3 - Reason_for_Exclusion_change
                                ],
            '% of Not Matched': [HICNO_change / rows3, first_name_change / rows3, last_name_change / rows3, 
                                Address_Line1_change / rows3, Address_Line2_change / rows3,
                                Address_Line3_change / rows3, Address_Line4_change / rows3,
                                Address_Line5_change / rows3, Address_Line6_change / rows3,
                                State_change / rows3, Zip_Code_change / rows3,
                                Gender_change / rows3, Birth_Date_change / rows3,
                                Date_of_Exclusion_change / rows3, Reason_for_Exclusion_change / rows3],
            'Compared Results': l                               
                                 
            })
    

#### Check Exclusion Reason Code for both Excel2 and IDR
    reasons_excel = {}
    reasons_idr = {}
    reasons_excel1 = {}
    reasons_idr1 = {}
    reasons_diff = {}
    reasons_compared = {}

    
    for row in excel2.index:
        v = excel2.loc[row, 'Reason for Exclusion']
        if v in reasons_excel.keys():
            reasons_excel[v] += 1
        else:
            reasons_excel[v] = 1
            reasons_idr[v] = 0
            reasons_excel1[v] = 0
            reasons_idr1[v] = 0
    
    for row in idr_excl_df.index:
        v = idr_excl_df.loc[row, 'Exclusion_Reason']
        if v in reasons_excel.keys():
            reasons_idr[v] += 1
        else:
            reasons_excel[v] = 0
            reasons_idr[v] = 1
            reasons_excel1[v] = 0
            reasons_idr1[v] = 0

    for row in excel1.index:
        v = excel1.loc[row, 'Reason for Exclusion']
        if v in reasons_excel1.keys():
            reasons_excel1[v] += 1
        else:
            reasons_excel[v] = 0
            reasons_idr[v] = 0
            reasons_excel1[v] = 1
            reasons_idr1[v] = 0
    
    for row in idr_excl_df1.index:
        v = idr_excl_df1.loc[row, 'Exclusion_Reason']
        if v in reasons_excel1.keys():
            reasons_idr1[v] += 1
        else:
            reasons_excel[v] = 0
            reasons_idr[v] = 0
            reasons_excel1[v] = 0
            reasons_idr1[v] = 1
            
   
    for key in reasons_idr:
        v1 = reasons_excel1[key]
        v2 = reasons_idr[key]
        reasons_diff[key] = v2 - v1

        if reasons_excel[key] == reasons_idr[key]:
            reasons_compared[key] = "TRUE"
        else:
            reasons_compared[key] = "FALSE"

    
 # reasons_compared   
            
#    exclusion_reasons_df = pandas.DataFrame({'Exclusion Reasons': list(reasons_excel.keys()),
#                                             '# Excel Current': list(reasons_excel.values()),
#                                             '# IDR Current': list(reasons_idr.values()),
#                                             '# Excel Previous': list(reasons_excel1.values()),
#                                             '# IDR Previous': list(reasons_idr1.values())})

    if model == 'NGACO':
        exclusion_reasons_df = pandas.DataFrame({'Exclusion Reasons': list(reasons_excel.keys()),
                                             '# Report Previous': list(reasons_excel1.values()),
                                             '# Current Month exclusion /DB(SQL)': list(reasons_diff.values()),
                                             '# Report Current': list(reasons_excel.values()),
                                             'Compared Results': list(reasons_compared.values())})
    else:
        exclusion_reasons_df = pandas.DataFrame({'Exclusion Reasons': list(reasons_excel.keys()),
                                             '# Report Previous Quarter': list(reasons_excel1.values()),
                                             '# Current Quarter exclusion /DB(SQL)': list(reasons_diff.values()),
                                             '# Report Current Quarter': list(reasons_excel.values()),
                                             'Compared Results': list(reasons_compared.values())})
    
#####################################################
#### Check IDR missing values and combine with Excel
#####################################################
                                             
### Drop "Performance_Year", "CurrentMonthDate" from IDR table
                                       
    idr_col_df = idr_excl_df.drop(["Performance_Year", "CurrentMonthDate"], axis=1)       
                                             
    idr_null = idr_col_df.isnull().sum()
    idr_nnull = idr_col_df.notnull().sum()
    idr_null_pct = list(range(0, len(idr_null)))
    
    idr_excel_match = []
    
    for i in range(len(df_nnull)):
        if abs(df_nnull[i] - idr_nnull[i]) < 10:
            idr_excel_match.append('TRUE')
        else:
            idr_excel_match.append('FALSE')
        
    
    for i in range(0, len(idr_null)):
        idr_null_pct[i] = idr_null[i] / (idr_null[i] + idr_nnull[i]) 

    freq = pandas.DataFrame({
          'Variables in File': dfcols,
          '# of Report Not Null': df_nnull.tolist(),
          '# of Report Null': df_null.tolist(),
          '% of Report Null': null_pct,
          '# of IDR Not Null': idr_nnull.tolist(),
          '# of IDR Null': idr_null.tolist(),
          '% of IDR Null': idr_null_pct,
          'Compared Results': idr_excel_match 
                })

#    if model == 'VTAPM':
#        report_info = Report 1-2" + "\n" + "Vermont All Payer Model" + "\n" + "Report on Excluded Beneficiaries" +  "Performance Year 2020"	+ "\n" +	
#                      "Cumulative Report Through Exclusion Run Apr 10, 2020"				
#    else:
 #               report_info = "Report 1-2" + "\n" + "Next Generation ACO Model" + "\n" + "Report on Excluded Beneficiaries" +  "Performance Year 2020"	+ "\n" +	
 #                     "Cumulative Report Through Exclusion Run June 10, 2020"			
    
 #   call object 
    di_1_2.collect_stats(freq, Excel_IDR_freq, exclusion_reasons_df, aco_name, sk) 
    
    csk = sk.get_csk()
    
    if model == 'NGACO':
        di_1_2.set_header(f"Report 1-2\nNext Generation ACO Model\nCurrent SK:  {csk}\nReport on Excluded Beneficiaries\nPerformance Year 2020" + excel0.iloc[2, 1])
    else:
        di_1_2.set_header(f"Report 1-2\nVermont All Payer Model\nCurrent SK:  {csk}\nReport on Excluded Beneficiaries\nPerformance Year 2020" + excel0.iloc[2, 1])
   
    with pandas.ExcelWriter(summary_file) as writer:
        
        # summary sheet
        summary.to_excel(writer, sheet_name = "Month_Over_Month_Summary", index=False, startrow=1)
    
        workbook = writer.book
        count_format = workbook.add_format({"num_format": "#,##0"})

        right_format = workbook.add_format({'align': 'right'})
        wrap_format = workbook.add_format()
        wrap_format.set_text_wrap()
 
       
   
        worksheet = writer.sheets["Month_Over_Month_Summary"]
        worksheet.merge_range('A1:B1', report_gen + "\n" + report_desc + f"\nCurrent SK: {csk}", wrap_format)
        worksheet.set_row(0, 60)
        worksheet.set_column("A:A", 60)
        worksheet.set_column("B:B", 18, count_format)

        count_format = workbook.add_format({"num_format": "#,##0"})
        pct_format = workbook.add_format({"num_format": "0.00%"})
    
        freq.to_excel(writer, sheet_name = "Summary_Of_Missing_Values", index=False, startrow=1)
        summary_sheet = writer.sheets["Summary_Of_Missing_Values"]
        summary_sheet.merge_range('A1:H1', report_gen + "\n" + report_desc + f"\nCurrent SK: {csk}", wrap_format)
        summary_sheet.set_row(0, 60)
        summary_sheet.set_column("A:A", 40)
        summary_sheet.set_column(1, 1, 18, count_format)
        summary_sheet.set_column(2, 2, 18, count_format)
        summary_sheet.set_column(3, 3, 18, pct_format)
        summary_sheet.set_column(4, 4, 18, count_format)
        summary_sheet.set_column(5, 5, 18, count_format)
        summary_sheet.set_column(6, 6, 18, pct_format)
        summary_sheet.set_column(7, 7, 18, right_format)

        Excel_IDR_freq.to_excel(writer, sheet_name = "Report_vs_IDR_Summary", index=False, startrow=1)
        Excel_IDR_freq_sheet = writer.sheets["Report_vs_IDR_Summary"]
        Excel_IDR_freq_sheet.merge_range('A1:E1', report_gen + "\n" + report_desc + f"\nCurrent SK: {csk}", wrap_format)
        Excel_IDR_freq_sheet.set_row(0, 60)
        Excel_IDR_freq_sheet.set_column("A:A", 40) 
        Excel_IDR_freq_sheet.set_column(1, 1, 18, count_format)
        Excel_IDR_freq_sheet.set_column(2, 2, 18, count_format)
        Excel_IDR_freq_sheet.set_column(3, 3, 18, pct_format)
        Excel_IDR_freq_sheet.set_column(4, 4, 18, right_format)
        
        exclusion_reasons_df.to_excel(writer, sheet_name = "Exclusion_Reason_Types_Check", index=False, startrow=1)
        exclusion_reasons_df_sheet = writer.sheets["Exclusion_Reason_Types_Check"]
        exclusion_reasons_df_sheet.merge_range('A1:E1', report_gen + "\n" + report_desc + f"\nCurrent SK: {csk}", wrap_format)
        exclusion_reasons_df_sheet.set_row(0, 60)
        exclusion_reasons_df_sheet.set_column("A:A", 60)
        exclusion_reasons_df_sheet.set_column(1, 1, 30, count_format)
        exclusion_reasons_df_sheet.set_column(2, 2, 35, count_format)
        exclusion_reasons_df_sheet.set_column(3, 3, 30, count_format)
        exclusion_reasons_df_sheet.set_column(4, 4, 20, right_format)
        


        
    with pandas.ExcelWriter(diff_file) as writer:
        only1.to_excel(writer, sheet_name = "Only in previous", index=False)
        only2.to_excel(writer, sheet_name = "Only in current", index=False)

