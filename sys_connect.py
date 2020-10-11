# -*- coding: utf-8 -*-
"""
Created on Sun Oct 11 15:29:23 2020

@author: heart
"""

import os

hostname = "xxx"

f = open(os.getenv("HOME") + "/.db_credential", "r")
username = f.readline().strip()
password = f.readline().strip()
f.close()
database = "xxx"

# IDR Teradata , username, password the same as above
IDR_host = ""

# Define file path for reading and exporting the validation reports
path = '/home/PRACTICE/Excel_File_Repository/'
outpath = '/home/output/'


