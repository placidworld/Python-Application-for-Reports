# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 12:57:08 2020

@author: heart
"""



import pandas
import ibm_db
import sqlalchemy

from PRACTICE.sys_connect import hostname, username, password, database, IDR_host, path, outpath

pandas.set_option('display.float_format', lambda x: '%.3f' % x)


class Report51Data:
    def __init__(self, header):
        self.header = header
        
    def get_header(self):
        return self.header

def compare_5_1_with_db2(infile2, outfile, outile2, model, ppo_name, program_code, year, di_5_1, sk):
    
    print("Comparing with db2 data ...")

    excel0 = pandas.read_excel(infile2, 
                               sheet_name=0, 
                               nrows = 1,
                               convert_float = True, 
                               header=0, 
                               index_col=None, 
                               dtype="object")

    py0 = excel0.iloc[0,11]
    performance_year = int(py0[-4:])
    
 # Reporting Month: Apr, 2020		Reporting Month Start Date: 04/01/2020		Reporting Month End Date: 04/30/2020			Performance Year: 2020		

    ### Read in report header and regroup for the validation report 
    report_month = excel0.iloc[0,4]
    report_month_start_date = excel0.iloc[0,6]
    report_month_end_date = excel0.iloc[0,8]
    report_performance_year = excel0.iloc[0,11]    
    
    
    excel2 = pandas.read_excel(infile2, 
                               sheet_name=0, 
                               nrows = 1,
                               convert_float = True, 
                               header=2, 
                               index_col=None, 
                               dtype="object")
    
    # remove Unnamed columns
    dfcols = list(excel2.columns)
    for c in excel2.columns:
        if isinstance(c, str) and "Unnamed" in c:
            dfcols.remove(c)
    excel2 = excel2.loc[:, dfcols]
    excel_total_beneficiaries = excel2.iloc[0, 1]
    excel_total_opt_out = excel2.iloc[0, 2]
    excel_total_claims_cnt = excel2.iloc[0, 4]
    excel_total_claims_amt = excel2.iloc[0, 5]

    
    # Retrieve beneficiaries from db2
    if model == 'PROGRAM2':
        sql = f"""
              SELECT *
              FROM ....
              AND C.CCO_ID = '{ppo_name}'
              AND LAST_DAY(ADD_MONTHS(CURRENT DATE, -1) ) BETWEEN A.EFCTV_DT AND A.TRMNTN_DT
              WITH UR;
       """
    else:
        sql = f"""
              SELECT COUNT(*) as COUNT
              FROM ...
              WHERE  A.CCO_PGM_CD = {program_code} 
              AND C.CCO_ID = '{ppo_name}'
             
              WITH UR;
        """      
    
    # use pconnect instead of connect - pconnect will keep connection open
    conn = ibm_db.pconnect(f"DATABASE={database};HOSTNAME={hostname};PORT=12003;PROTOCOL=TCPIP;UID={username};PWD={password}",'','')
    stmt = ibm_db.prepare(conn,sql)
    
    print("Running SQL query ...")
    ibm_db.execute(stmt)
    print("Running SQL query ... done")
    
    result_dict = ibm_db.fetch_assoc(stmt)
    
    db2_data = []
    while result_dict:
        db2_data.append(result_dict)
        result_dict = ibm_db.fetch_assoc(stmt)

        
    db2_df = pandas.DataFrame(db2_data)
    db2_total_beneficiaries = db2_df.loc[0, 'COUNT']
    
    if excel_total_beneficiaries != db2_total_beneficiaries:
        print(f"Total count of beneficiaries does not match: {excel_total_beneficiaries} (excel) != {db2_total_beneficiaries} (db2)")
        bene_match = "Not Match"
    else:
        print("Total count of beneficiaries matches.")
        bene_match = "Match"
        
    # Retrieve total beneficiary opt-out
    sql = f"""
              SELECT COUNT(*) as COUNT
              FROM ...
              WITH UR;
        """
            

    stmt = ibm_db.prepare(conn,sql)
    
    print("Running SQL query ...")
    ibm_db.execute(stmt)
    print("Running SQL query ... done")
    
    result_dict = ibm_db.fetch_assoc(stmt)
    
    db2_data = []
    while result_dict:
        db2_data.append(result_dict)
        result_dict = ibm_db.fetch_assoc(stmt)
        
    db2_df = pandas.DataFrame(db2_data)
    db2_total_opt_out = db2_df.loc[0, 'COUNT']
    
    if excel_total_opt_out != db2_total_opt_out:
        print(f"Total beneficiary opt-out does not match: {excel_total_opt_out} (excel) != {db2_total_opt_out} (db2)")
        opt_out_match = "Not Match"
    else:
        print("Total beneficiary opt-out matches.")
        opt_out_match = "Match"


### Substance Abuse ####
### Substance Abuse "Count of SA Claims" validation
    excel_sa = pandas.read_excel(infile2, 
                               sheet_name=1, 
                               nrows = 1,
                               convert_float = True, 
                               header=2, 
                               index_col=None, 
                               dtype="object")
    
    # remove Unnamed columns
    dfcols = list(excel_sa.columns)
    for c in excel_sa.columns:
        if isinstance(c, str) and "Unnamed" in c:
            dfcols.remove(c)
    excel_sa = excel_sa.loc[:, dfcols]
    excel_total_bene_SA = excel_sa.iloc[0, 5]
    excel_optin_SA_CLM = excel_sa.iloc[0, 7]
    excel_SUPRSD_SA_CLM = excel_sa.iloc[0, 11]
    

    ######################################################################################
    # Retrieve data from DB2 for Total Count of Beneficiaries Opted-In to SA Data Sharing
    sql = f"""
          SELECT COUNT(1) AS COUNT
          FROM ...
          with UR;   
        """         

    conn = ibm_db.pconnect(f"DATABASE={database};HOSTNAME={hostname};PORT=12003;PROTOCOL=TCPIP;UID={username};PWD={password}",'','')
    stmt = ibm_db.prepare(conn,sql)
    
    print("Running SQL query ...")
    ibm_db.execute(stmt)
    print("Running SQL query ... done")
    
    result_dict = ibm_db.fetch_assoc(stmt)
    
    db2_sa = []
    while result_dict:
        db2_sa.append(result_dict)
        result_dict = ibm_db.fetch_assoc(stmt)

        
    db2_sa_df = pandas.DataFrame(db2_sa)
    db2_total_bene_SA = db2_sa_df.loc[0, 'COUNT']
    
    if excel_total_bene_SA != db2_total_bene_SA:
        print(f"Total count of beneficiaries SA does not match: {excel_total_bene_SA} (excel) != {db2_total_bene_SA} (db2)")
        bene_SA_match = "Not Match"
    else:
        print("Total count of beneficiaries SA matches.")
        bene_SA_match = "Match"
    


    
#### Retrieve the following from Teradata 
#         “Count of SA Claims” 
#         “Count of SA Opt-in Claims”
#         “Count of SA Claims Suppressed”

    if model == 'PROGRAM2':
        sql = f"""
               SELECT TOTAL_SA_CLM_COUNT, 
                      SA_CLM_OPT_IN_COUNT,
                      SA_SUPRSD_CLM_COUNT
               FROM ...
               WHERE PRVDR_ACO_ID = '{ppo_name}'
        """
    else:    
        sql = f"""
               SELECT TOTAL_SA_CLM_COUNT, 
                      SA_CLM_OPT_IN_COUNT,
                      SA_SUPRSD_CLM_COUNT
               FROM ...
               WHERE PRVDR_ACO_ID = '{ppo_name}'
        """

    ### Connect to Teradata
    engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
    idr_sa_df = pandas.read_sql(sql,con=engine)
     
    
    idr_SA_CLM = idr_sa_df.loc[0, 'TOTAL_SA_CLM_COUNT']
    idr_optin_SA_CLM = idr_sa_df.loc[0, 'SA_CLM_OPT_IN_COUNT']
    idr_SUPRSD_SA_CLM = idr_sa_df.loc[0, 'SA_SUPRSD_CLM_COUNT']    

    excel_SA_CLM = excel_sa.iloc[0, 1]
    excel_optin_SA_CLM = excel_sa.iloc[0, 7]
    excel_SUPRSD_SA_CLM = excel_sa.iloc[0, 11]


    if excel_SA_CLM != idr_SA_CLM:
        print(f"Total SA Claims do not match: {excel_SA_CLM} (excel) != {idr_SA_CLM} (idr)")
        SA_CLM_match = "Not Match"
    else:
        print("Total SA Claims match.")
        SA_CLM_match = "Match"

    if excel_optin_SA_CLM != idr_optin_SA_CLM:
        print(f"Count of SA Opt-in Claims do not match: {excel_optin_SA_CLM} (excel) != {idr_optin_SA_CLM} (idr)")
        OPTIN_SA_CLM_match = "Not Match"
    else:
        print("Count of SA Opt-in Claims match.")
        OPTIN_SA_CLM_match = "Match"

    if excel_SUPRSD_SA_CLM != idr_SUPRSD_SA_CLM:
        print(f"Count of SA Claims Suppressed do not match: {excel_SUPRSD_SA_CLM} (excel) != {idr_SUPRSD_SA_CLM} (idr)")
        SUPRSD_SA_CLM_match = "Not Match"
    else:
        print("Count of SA Claims Suppressed match.")
        SUPRSD_SA_CLM_match = "Match"
              
########################
### Exclusions Worksheet
    excel_excl = pandas.read_excel(infile2, 
                               sheet_name=3, 
                               nrows = 1,
                               convert_float = True, 
                               header=2, 
                               index_col=None, 
                               dtype="object")
    
    # remove Unnamed columns
    dfcols = list(excel_excl.columns)
    for c in excel_excl.columns:
        if isinstance(c, str) and "Unnamed" in c:
            dfcols.remove(c)
    excel_excl = excel_excl.loc[:, dfcols]

### MC - Transition to Medicare Advantage (MA)	
### MS - Medicare as Secondary Payer	
### DD - Date of Death occurs prior to the start of the PY	
### RP - Beneficiary aligned to another Program	
### EP - Beneficiary aligned to another Program	
### AB - Loss of Part A or Part B		
### CO - Moved outside of the ACO’s ESA or from the July of AY2 county to a non ACO ESA county	
### OU - Not a resident of US or its territories		
### EM - Has at least one PQEM and 1) No PQEM with ACO Participant or 2) No PQEM with Participant within ESA		
### NV - Eligibility cannot be verified		
### Total Count of  Excluded Beneficiaries

    exclusion_code = ['MC', 'MS', 'DD', 'RP', 'EP', 'AB', 'CO', 'OU', 'EM', 'NV', 'To']
    
    exclusion_desc = {}
    excel_exclusion_values = {}
    
    # initialize dictionaries
    for key in exclusion_code:
        exclusion_desc[key] = key + " - "
        excel_exclusion_values[key] = 0

    # read values from excel table excep the first column
    for col in excel_excl.columns[1:]:
        key = col[0:2]
        if key not in exclusion_code:
            print(f"Unknown exclusion key {key} from excel sheet.")
            continue
        
        exclusion_desc[key] = col
        excel_exclusion_values[key] = excel_excl.loc[0, col]


    csk = sk.get_csk()
    psk = sk.get_psk()

# =============================================================================
#     if model == "PROGRAM1":
#         csk = sk_dic["CN51"]
#         psk = sk_dic["PN51"]
#     else:
#         csk = sk_dic["CV51"]
#         psk = sk_dic["PV51"]
# =============================================================================
       
### Read data from IDR
    if model == 'PROGRAM1':
        sql = f"""        
                  SELECT *
                 FROM 	..
        """      
    else:
        sql = f"""
               SElECT  *
               from    ...
                """

    
    engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
    idr_excl_df = pandas.read_sql(sql,con=engine)
    
    # dictionary holding key value pair from database
    idr_exclusion_values = {}
    for key in exclusion_code:
        idr_exclusion_values[key] = 0
            
    # read values from databaes table
    for i in idr_excl_df.index:
        key = idr_excl_df.loc[i, 'EXCL_CD'].strip()
        value = idr_excl_df.loc[i, 'EXCL_BENE_CT']
        
        idr_exclusion_values[key] = value
        idr_exclusion_values['To'] += value

    # compare excel data with database
    idr_match = {}
    
    for key in exclusion_code:
        if key not in excel_exclusion_values or key not in idr_exclusion_values:
            idr_match[key] = "Missing"
        elif excel_exclusion_values[key] != idr_exclusion_values[key]:
            idr_match[key] = "Not Match"
        else:
            idr_match[key] = "Match"





##############################################################################
### Exclusions DB2
    if model == 'PROGRAM1':
        sql = f"""        
                   SELECT *
                   FROM ...
                 WITH UR;
        """      
    else:
        sql = f"""
               SELECT *
               FROM ...
              WITH UR;                
        """

    
    conn = ibm_db.pconnect(f"DATABASE={database};HOSTNAME={hostname};PORT=12003;PROTOCOL=TCPIP;UID={username};PWD={password}",'','')
    stmt = ibm_db.prepare(conn,sql)
        
    print("Running SQL query ...")
    ibm_db.execute(stmt)
    print("Running SQL query ... done")
    
    result_dict = ibm_db.fetch_assoc(stmt)
    
    db2_excl = []
    while result_dict:
        db2_excl.append(result_dict)
        result_dict = ibm_db.fetch_assoc(stmt)
        
    db2_excl_df = pandas.DataFrame(db2_excl)
    
    # dictionary holding key value pair from database
    db2_exclusion_values = {}
    for key in exclusion_code:
        db2_exclusion_values[key] = 0
            
    # read values from databaes table
    for i in db2_excl_df.index:
        key = db2_excl_df.loc[i, 'EXCL_CD'].strip()
        value = db2_excl_df.loc[i, 'EXCL_BENE_CT']
        
        db2_exclusion_values[key] = value
        db2_exclusion_values['To'] += value

    # compare excel data with database
    db2_match = {}
    
    for key in exclusion_code:
        if key not in excel_exclusion_values or key not in db2_exclusion_values:
            db2_match[key] = "Missing"
        elif excel_exclusion_values[key] != db2_exclusion_values[key]:
            db2_match[key] = "Not Match"
        else:
            db2_match[key] = "Match"
 

####################################
### Read in CCLF Run date from IDR
####################################

    if model == 'PROGRAM1':
        sql = f"""
              
              ;

        """
    else:    
        sql = f"""
              
             ;
             """
             

    engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
    CCLF_rundate_data = pandas.read_sql(sql,con=engine)

    CCLF_rundate_df = pandas.DataFrame(CCLF_rundate_data)
    CCLF_run_dt = CCLF_rundate_df.loc[0, 'CCLF_RUN_DT']
    CCLF_dt = CCLF_run_dt.strftime('%m/%d/%Y')
        
    CCLF_rundate = "CCLF Run Date:  " +  CCLF_dt
    
    
##########################################
### Read in SHRU Execution date from IDR
#########################################
    sql = f"""
              SELECT  MAX(CAST (meta_sk /1000 AS DATE)) AS SHRU_RUN_DATE
              FROM  ...
              ;
        """
             
    SHRU_rundate_data = []
    conn = None
    

    engine = sqlalchemy.create_engine('teradatasql://{}:{}@{}/?logmech=LDAP'.format(username,password,IDR_host))
    SHRU_rundate_data = pandas.read_sql(sql,con=engine)
     

    SHRU_rundate_df = pandas.DataFrame(SHRU_rundate_data)
    SHRU_run_dt = SHRU_rundate_df.loc[0, 'SHRU_RUN_DATE']
    SHRU_dt = SHRU_run_dt.strftime('%m/%d/%Y')
        
    SHRU_rundate = "Data Sharing Preference Load Date:  " +  SHRU_dt
    
#########################################
### Read in DPREF Date from DB2 
########################################

    # Retrieve beneficiaries from db2
    if model == 'PROGRAM2':
        sql = f"""
              SELECT *
              FROM ... 
              WITH UR;
       """
    else:
        sql = f"""
              SELECT *
              FROM >..
              WITH UR;
        """      
    
    DPREF_rundate_data = []
    conn = None
    
    
    try:
        conn = ibm_db.pconnect(f"DATABASE={database};HOSTNAME={hostname};PORT=12003;PROTOCOL=TCPIP;UID={username};PWD={password}",'','')
        stmt = ibm_db.prepare(conn,sql)
        
        print("Running SQL query ...")
        ibm_db.execute(stmt)
        print("Running SQL query ... done")
        
        result_dict = ibm_db.fetch_assoc(stmt)
        
        while result_dict:
            DPREF_rundate_data.append(result_dict)
            result_dict = ibm_db.fetch_assoc(stmt)

    except Exception as e:
        print(str(e))
        
    DPREF_rundate_df = pandas.DataFrame(DPREF_rundate_data)
    DPREF_run_dt = DPREF_rundate_df.loc[0, 'DPREF_RUN_DATE']
    DPREF_dt = DPREF_run_dt.strftime('%m/%d/%Y')
        
    DPREF_rundate = "DPREF Run Date:  " +  DPREF_dt


######################################################
#### PBP AIPBP worksheet data
######################################################
    excel_pbp = pandas.read_excel(infile2, 
                               sheet_name=2, 
                               nrows = 1,
                               convert_float = True, 
                               header=2, 
                               index_col=None, 
                               dtype="object")
    
    # remove Unnamed columns
    dfcols = list(excel_pbp.columns)
    for c in excel_pbp.columns:
        if isinstance(c, str) and "Unnamed" in c:
            dfcols.remove(c)
    excel_pbp = excel_pbp.loc[:, dfcols]
    excel_bene_w_PBP = excel_pbp.iloc[0, 1]


#########################################
### Read in Audit Data from DB2
########################################

    # Retrieve beneficiaries from db2
    if model == 'PROGRAM2':
        sql = f"""

       """
    else:
        sql = f"""
 """      
        
    audit_data = []
# excel_exclusion_values   
    excel_total_exclusion_cnt = excel_exclusion_values['To']
    
    conn = None
    
    
    try:
        conn = ibm_db.pconnect(f"DATABASE={database};HOSTNAME={hostname};PORT=12003;PROTOCOL=TCPIP;UID={username};PWD={password}",'','')
        stmt = ibm_db.prepare(conn,sql)
        
        print("Running SQL query ...")
        ibm_db.execute(stmt)
        print("Running SQL query ... done")
        
        result_dict = ibm_db.fetch_assoc(stmt)
        
        while result_dict:
            audit_data.append(result_dict)
            result_dict = ibm_db.fetch_assoc(stmt)

    except Exception as e:
        print(str(e))
        
    audit_df = pandas.DataFrame(audit_data)
    
    for row in audit_df.index:
        if 'BENE_CNT' in audit_df.loc[row, 'CLMN_NAME']:
            audit_total_beneficiaries = audit_df.loc[row, 'STATS_CNT']
        elif 'BENEEXCL_COUNT' in audit_df.loc[row, 'CLMN_NAME']:
            audit_total_exclusion_cnt = audit_df.loc[row, 'STATS_CNT']
        elif 'PBP_CLM_CNT' in audit_df.loc[row, 'CLMN_NAME']:
            audit_bene_w_PBP = audit_df.loc[row, 'STATS_CNT']
        elif 'CLM_CNT' in audit_df.loc[row, 'CLMN_NAME']:
            audit_total_claims_cnt = audit_df.loc[row, 'STATS_CNT']
        elif 'TOL_CLM_AMT' in audit_df.loc[row, 'CLMN_NAME']:
            audit_total_claims_amt = audit_df.loc[row, 'STATS_CNT']    

    audit_update_date = audit_df.loc[0, 'REC_UPDT_DT']
    audit_run_dt = 'Audit Table Updated Date:  ' + audit_update_date
            
#    DPREF_run_dt = DPREF_rundate_df.loc[0, 'DPREF_RUN_DATE']
#    DPREF_dt = DPREF_run_dt.strftime('%m/%d/%Y')
        
#    DPREF_rundate = "DPREF Run Date:  " +  DPREF_dt

    if excel_total_beneficiaries != audit_total_beneficiaries:
        total_beneficiaries_match = "Not Match"
    else:
        total_beneficiaries_match = "Match"


    if excel_total_claims_cnt != audit_total_claims_cnt:
        total_claims_cnt_match = "Not Match"
    else:
        total_claims_cnt_match = "Match"

    if int(excel_total_claims_amt) != int(audit_total_claims_amt):
        total_claims_amt_match = "Not Match"
    else:
        total_claims_amt_match = "Match"

    if excel_bene_w_PBP != audit_bene_w_PBP:
        bene_w_PBP_match = "Not Match"
    else:
        bene_w_PBP_match = "Match"

    if excel_total_exclusion_cnt != audit_total_exclusion_cnt:
        total_exclusion_cnt_match = "Not Match"
    else:
        total_exclusion_cnt_match = "Match"


#    with pandas.ExcelWriter(outfile) as writer:

    df = pandas.DataFrame({'Type' : ["Total Count of Beneficiaries", 'Total Count of Beneficiaries Opted-Out of Medical Data Sharing'],
                           'Report': [excel_total_beneficiaries, excel_total_opt_out],
                           'DB2': [db2_total_beneficiaries, db2_total_opt_out],
                           'Compared Results': [bene_match, opt_out_match]})

    df_sa = pandas.DataFrame({'Type' : ["Total Count of Beneficiaries Opted-In to SA Data Sharing", "Total SA Claims",
                                        "Count of SA Opt-in Claims", "Count of SA Claims Suppressed"],
                              'Report': [excel_total_bene_SA, excel_SA_CLM, excel_optin_SA_CLM, excel_SUPRSD_SA_CLM],
                              'DataBase Values': [db2_total_bene_SA, idr_SA_CLM, idr_optin_SA_CLM, idr_SUPRSD_SA_CLM],
                              'Compared Results': [bene_SA_match, SA_CLM_match, OPTIN_SA_CLM_match, SUPRSD_SA_CLM_match]})

    df_excl = pandas.DataFrame({'Bene Exclusion Reason Type' : [v for v in exclusion_desc.values()],
                                'Report': [v for v in excel_exclusion_values.values()],
                                'IDR'  : [v for v in idr_exclusion_values.values()],
                                'Report VS IDR Results': [v for v in idr_match.values()],
                                'DB2'  : [v for v in db2_exclusion_values.values()],
                                'Report VS DB2 Results': [v for v in db2_match.values()]
                               })

    df_audit = pandas.DataFrame({'Type': ["Total Count of Beneficiaries", "Count of Total Claims (M+SA)",
                                          "Total Claims Amount (M+SA)", "Count of Beneficiaries with PBP/AIPBP",
                                          "Total Count of Excluded Beneficiaries"],
                                'Report': [excel_total_beneficiaries, excel_total_claims_cnt,
                                          excel_total_claims_amt, excel_bene_w_PBP,
                                          excel_total_exclusion_cnt],
                                'Audit Table': [audit_total_beneficiaries, audit_total_claims_cnt,
                                          audit_total_claims_amt, audit_bene_w_PBP,
                                          audit_total_exclusion_cnt],
                                'Compared Results': [total_beneficiaries_match, total_claims_cnt_match,
                                          total_claims_amt_match, bene_w_PBP_match,
                                          total_exclusion_cnt_match]})

 #   call object 
    di_5_1.collect_stats(df, df_sa, df_excl, df_audit, ppo_name)
    
    writer = pandas.ExcelWriter(outfile)

    df.to_excel(writer, sheet_name='Total Statistics', index=None, startrow=1)
    df_sa.to_excel(writer, sheet_name='Substance Abuse', index=None, startrow=1)
    df_excl.to_excel(writer, sheet_name='Exclusions', index=None, startrow=1)
    df_audit.to_excel(writer, sheet_name='Audit Table Validation', index=None, startrow=1)
    #df_bene_sa.to_excel(writer, sheet_name='Substance Abuse', index=None)
    
        
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
    
    header = "{}\n{}\n{}\n{}\nCurrent SK used: {}\nPrevious SK used: {}\n{}\n{}\n{}\n{}".format(
        report_month,
        report_month_start_date,
        report_month_end_date,
        report_performance_year,
        csk,
        psk,
        CCLF_rundate,
        SHRU_rundate,
        DPREF_rundate,
        audit_run_dt)
    
    di_5_1.set_header(header)
    report_5_1_data = Report51Data(header)
                        
    for worksheet in writer.sheets.values():
                
#        worksheet.write(0, 0, report_month + '\n' + report_month_start_date + '\n' + 
#                        report_month_end_date + '\n' + report_performance_year, bold_format)
#        worksheet.write(0, 1, report_month_start_date)
#        worksheet.write(0, , report_month_end_date)
#        worksheet.write(0, 4, report_performance_year)
        worksheet.merge_range('A1:D1', header, wrap_format)
        worksheet.set_row(0, 120)
        worksheet.set_column("A:A", 68)
        worksheet.set_column("B:B", 30, count_format)
        worksheet.set_column(2, 2, 30, count_format)
        worksheet.set_column(3, 3, 30, right_align_format)
        worksheet.set_column(4, 4, 30, count_format)
        worksheet.set_column(5, 5, 30, right_align_format)
        worksheet.set_column(6, 6, 30, count_format)
        worksheet.set_column(7, 7, 30, count_format)
        worksheet.set_column(8, 8, 30, right_align_format)

    writer.save()
    return report_5_1_data
 
