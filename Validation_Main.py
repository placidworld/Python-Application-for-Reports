# -*- coding: utf-8 -*-
"""
Created on Tue Oct  6 19:15:33 2020

@author: heart
"""


import pandas
import os
import re


#from PRACTICE.compare_excel import compare_excel

from PRACTICE.compare_phi_1_1 import compare_phi_1_1
from PRACTICE.compare_phi import compare_phi

from PRACTICE.data_inconsistency_1_2 import DataInconsist_1_2

from PRACTICE.compare_excel_2_2 import compare_excel_2_2
from PRACTICE.compare_excel_2_3 import compare_excel_2_3


from PRACTICE.PROGRAM1_5_1_PPO_Summary import PROGRAM1_5_1_PPO_Sum
from PRACTICE.compare_excel_5_1 import compare_excel_5_1
from PRACTICE.compare_excel_5_1M import compare_excel_5_1M
from PRACTICE.compare_5_1_with_db2 import compare_5_1_with_db2
from PRACTICE.compare_5_1M_with_db2 import compare_5_1M_with_db2

from PRACTICE.compare_excel import compare_excel



from PRACTICE.data_inconsistency_5_1 import DataInconsist_5_1
from PRACTICE.compare_6_1 import Compare_6_1

#from PRACTICE.compare_excel_6_1 import compare_excel_6_1

from PRACTICE.compare_excel_6_2 import compare_excel_6_2

from PRACTICE.compare_phi_6_3 import compare_phi_6_3
from PRACTICE.data_inconsistency_6_3 import DataInconsist_6_3

from PRACTICE.compare_phi_6_4 import compare_phi_6_4
from PRACTICE.validate_6_4_category import Category_6_4

from PRACTICE.compare_excel_PROGRAM3 import compare_excel_PROGRAM3

from PRACTICE.PROGRAM3_Report_VS_Tera import PROGRAM3_Report_VS_Tera

from PRACTICE.PROGRAM1_6_3_Report_Generated import PROGRAM1_6_3_Report_Generated

from PRACTICE.PROGRAM1_PROGRAM2_SK import PROGRAM1_PROGRAM2_SK

# Email sent out throughn outlook 
from PRACTICE.Email_From_Bears import send_email

# System setup - input file location, output file location, username, password
from PRACTICE.sys_connect import path, outpath

# inititalize those values as global
di_1_2 = None
di_5_1 = None
di_6_3 = None

report_generated =None

# Default Change % for total counts
tot_change = 2

# Default acceptable % for null values in one column
null_change = 3

sk = PROGRAM1_PROGRAM2_SK()

validate_6_4_category = Category_6_4()


pandas.set_option('display.float_format', lambda x: '%.3f' % x)


# define parameters for dynamic running
def do_process(infile1, infile2, model, report_num, PPO_name, error_threshold):

    # generate output file name
    # this can be as many as needed
    fname = infile2.split(".")
    fname1 = fname[0]

    fname[0] = fname1 + "_Previous_Month_Comparison_results"
    outfile = outpath + ".".join(fname)

    fname[0] = fname1 +  "_DB_validation_result"
    outfile2 = outpath + ".".join(fname)  

    fname[0] = fname1 +  "_Differences"
    outfile3 = outpath + ".".join(fname)

    fname[0] = fname1 + "_summary_all_PPO"
    outfile4 = outpath + ".".join(fname) 

    fname[0] = fname1 + "_Excel_Tera_Changes"
    outfile5 = outpath + ".".join(fname) 

    fname[0] = fname1 + "_Report_vs_Tera_Results"
    outfile6 = outpath + ".".join(fname) 

    fname[0] = fname1 + "_Report_Validation"
    outfile7 = outpath + ".".join(fname) 

    fname[0] = fname1 + "_Previous_Year_Comparison_and_DB_Validation"
    outfile8 = outpath + ".".join(fname)

   
    path2 = path + infile2
   

    #
    # current month only processing
    #

    # Special treatment for this particular report due to it's nature, need check data inconsistency, use special SK etc
    if report_num == "5_1":
        if model == "PROGRAM2":
            r51data =compare_5_1_with_db2(path2, outfile2, outfile4, model, PPO_name, 53, 2020, di_5_1, sk)
        else:
            r51data = compare_5_1_with_db2(path2, outfile2, outfile4, model, PPO_name, 21, 2020, di_5_1, sk)
    elif report_num == "5_1M":
        if model == "PROGRAM2":
            r51mdata = compare_5_1M_with_db2(path2, outfile2, outfile4, model, 53, 2020, di_5_1, sk)
        else:
            r51mdata = compare_5_1M_with_db2(path2, outfile2, outfile4, model, 21, 2020, di_5_1, sk)
            
    if infile1 is None:
        return

    #
    # compare with previous month
    #

    path1 = path + infile1

    if report_num == "6_3":
        compare_phi_6_3(path1, path2, outfile2, outfile3, 5, 2, model, report_generated, sk, PPO_name, di_6_3, tot_change, null_change)
    elif report_num == "6_4":
        compare_phi_6_4(path1, path2, outfile, sheet=0, header=5, sk=sk)
        validate_6_4_category.validate_6_4_category(path2)        
    elif report_num == '1-2' or report_num == '1_2':
        compare_phi(path1, path2, outfile, outfile3, outfile5, model, PPO_name, sheet=0, header=3, sk=sk, di_1_2=di_1_2, error_threshold=error_threshold)
    elif report_num == "2_2" or report_num == "2-2":
        compare_excel_2_2(path1, path2, outfile7)        
    elif report_num == "2_3" or report_num == "2-3":
        compare_excel_2_3(path1, path2, outfile7)
    elif model == "PROGRAM3":
        compare_excel_PROGRAM3(path1, path2, outfile)
        PROGRAM3_Report_VS_Tera(path2, outfile6)

    elif report_num == "5_1":
        compare_excel_5_1(path1, path2, outfile, r51data=r51data)

    elif report_num == "5_1M":
        compare_excel_5_1M(path1, path2, outfile, r51data=r51mdata)        

    elif report_num == "6_1":
        compare_6_1.compare_6_1(path1, path2, model, PPO_name)
        compare_excel(path1, path2, outfile)

    elif report_num == "6_2":
        compare_excel_6_2(path1, path2, outfile, sk)        

    elif report_num == "1_1":
        compare_phi_1_1(path1, path2, outfile3, outfile8, model, PPO_name, sk=sk, error_threshold=error_threshold)
    else:
        compare_excel(path1, path2, outfile)
            

if __name__ == '__main__':

    error_threshold = None

    while True:
        model = input("Enter Program/Model, PROGRAM1/PROGRAM2/PROGRAM3:  ")
        if model in ['PROGRAM1', 'PROGRAM2', 'PROGRAM3']:
            break    

    if model == "PROGRAM3":
        report_num = "CCLF"
    else:
        report_num = input("Enter report number(1_1, 1_2, 1_2, 2_2, 2_3, 5_1, 5_1M, 6_1, 6_2, 6_3, 6_4):  ")
        while True:
            if report_num in ['1_1', '1_2', '2_2', '2_3', '5_1', '5_1M', '6_1', '6_2', '6_3', '6_4']:
                break           

    if report_num == "5_1" or report_num == "5_1M":
        di_5_1 = DataInconsist_5_1()
    elif report_num == "1_2":
        di_1_2 = DataInconsist_1_2()
      
    if report_num == "1_1" or report_num == "2_2" or report_num == "2_3":
        while True:
            current_date = input("Enter report year current (YYYY): ")
            if re.match(r'^20[12][0-9]$', current_date):
                break

        while True:
            previous_date = input("Enter report year previous (YYYY): ") 
            if re.match(r'^20[12][0-9]$', previous_date):
                break

        PPO_name = input("Enter PPO ID or ALL for all PPOs: ")

        if PPO_name == "ALL":
            PPO_name = "[_A-Za-z0-9]+"

        pattern = re.compile(r"^" + model + "_" + report_num + "_(" + PPO_name + ")_Annual_(" + previous_date + "|" + current_date + ")\.")   
    else:            
        while True:
            current_date = input("Enter report month current (YYYYMM): ")
            if re.match(r'^20[12][0-9][01][0-9]$', current_date):
                break            

        while True:
            previous_date = input("Enter report month previous (YYYYMM): ") 
            if re.match(r'^20[12][0-9][01][0-9]$', previous_date):
                break

        if model == 'PROGRAM1' and report_num == '5_1':
            PROGRAM1_5_1_PPO_Sum(previous_date, current_date)

        if model == "PROGRAM1" and report_num == "6_3":
            report_generated = PROGRAM1_6_3_Report_Generated()
            report_generated.collect_PPO_names(model, report_num, current_date)
            report_generated.query_PPO()
            di_6_3 = DataInconsist_6_3()

            
        if model == "PROGRAM3":
            PPO_name = "ALL"

        elif report_num == "5_1M":
            PPO_name = 'Monthly_CCLF_Management_Report'
        else:
            PPO_name = input("Enter PPO you want to validate or ALL for all PPOs: ")

        if PPO_name == "ALL":
            PPO_name = "[_A-Za-z0-9]+"

        elif report_num == '6_1':
            global compare_6_1
            compare_6_1 = Compare_6_1()

        # set up error threshold for report 6-3. Default value is inside the program, user can input different value            
        elif report_num == "6_3":
            val1 = input("Enter Total Records Change Current vs. Previous Threshold value in percentage [2]: ")
            if val1.strip() == "":
                tot_change = 2
            else:
                tot_change = int(val1.strip())

            val2 = input("Enter % of Null change Threshold value[3]: ")

            if val2.strip() == "":
                null_change = 3
            else:
                null_change = int(val2.strip())

                
        pattern = re.compile(r"^" + model + "_" + report_num + "_ ?(" + PPO_name + ")_(" + previous_date + "|" + current_date + ")\d\d\.")

    if report_num == "1_2" or report_num == "1_1":
        line = input("Enter Error Threshold value in percentage [2]: ")
        if line.strip() == "":
            error_threshold = 2
        else:
            error_threshold = int(line.strip())

    ### Funcation to define whether to Send Email or Not. Default is N which means do nothing, just click ENTER   
    send_notification_email = input("Send email [Y/N]: ")

    ### Pick the most recent SK value from SQL for business
    if model != "PROGRAM3":
        sk.query(model)

    ### Check whether the files to be validated exist or not
    ### use os.listdir 
    files = os.listdir(path)
    files.sort()

    prev_files = {}

    curr_files = {}

    for fname in files:
        m = pattern.match(fname)
        if not m:
            continue

        print("matched:" + fname)

        PPO = m.group(1)

        date = m.group(2)

        if date == previous_date:
            prev_files[PPO] = fname
        else:
            curr_files[PPO] = fname


    for PPO in curr_files:
        infile2 = curr_files[PPO]
        if PPO in prev_files:
            infile1 = prev_files[PPO]
        else:
            infile1 = None

        do_process(infile1, infile2, model, report_num, PPO, error_threshold)

    ### Create 5-1 summary output file name based on patterns
    if report_num == "5_1" and infile2 is not None:
        m = re.match(r".*_([0-9]+).xlsx", infile2)
        di_5_1.write(outpath + model + "_5_1_" + m.group(1) + "_summary_all_PPO.xlsx")
    elif report_num == "1_2" and infile2 is not None:
        m = re.match(r".*_([0-9]+).xlsx", infile2)

        di_1_2.write(outpath + model + "_1_2_" + m.group(1) + "_summary_all_PPO.xlsx")

        

    if report_num == "6_1" and infile2 is not None:
        audit_file = outpath + model + '_6_1_Audit_Summary_' + current_date + '.xlsx'
        compare_6_1.compare_6_1_save(audit_file)
   
    if model == "PROGRAM1" and report_num == "6_3":
        report_generated.gen_report(outpath + "/" + model + "_" + report_num + "_PPOs_Summary_" + current_date + ".xlsx")
        di_6_3.write(outpath + "/" + model + "_" + report_num + "_Comparison_Summary_" + current_date + ".xlsx")

    if report_num == "6_4" and infile2 is not None:
        validate_6_4_category.save(outpath + "/" + model + "_" + report_num + "_Entitlement_Category_Code_validation_" + current_date + ".xlsx", sk)


    ### Create dynamic email subject line based on repor
    subject_line = None
    rpt_num = report_num.replace("_", "-")  


    if send_notification_email != "n" and send_notification_email != "N":
        if model == "PROGRAM3":
            subject_line = model + " CCLF Monthly" + " " + "as of " + current_date + " validation is done!"
        else:
            subject_line = model + " " + rpt_num + " " + current_date + " validation is done!"

        if report_num == '6_1':
            send_email('Abitha.Padmanabhan@Bears.com,alka.kunwar@Bears.com,emily.li@Bears.com',
                subject_line,
                "Dear Prod Ops Team,\n\nPlease review and let us know if you have any questions or concerns.\n\nData Automation Team",
                audit_file
                )
        else:
            send_email('Abitha.Padmanabhan@Bears.com,alka.kunwar@Bears.com,emily.li@Bears.com',
                subject_line,
                "Dear Prod Ops Team,\n\nPlease review and let us know if you have any questions or concerns.\n\nData Automation Team")


