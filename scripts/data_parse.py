#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon May 10 16:25:21 2021

@author: virati
Filter through patient list for Lung Mets Specifically

"""

import pandas as pd
import openpyxl as opyx
from openpyxl import Workbook
from collections import defaultdict



''' Auto nested dictionary '''
def nestdict():
    return defaultdict(nestdict)


#%%
wb = opyx.load_workbook('../data/brain_mets.xlsx')
#pd.read_excel('../data/brain_mets.xlsx')
ws = wb.active
#%%

charts = []
dcharts = nestdict()

for row in ws.iter_rows(min_row = 4, max_col = 5, max_row=4432):#ws['B4':'B4432']:
    pt = row[1].value
    
    if pt != None: dcharts[pt]['Dx'] = row[4].value; #print(pt)
#%%
print(str(len(dcharts)) + ' total patients.')

#%%
import xlsxwriter
out_frame = pd.DataFrame(dcharts).T

writer = pd.ExcelWriter('pt_list.xlsx', engine='xlsxwriter')
out_frame.to_excel(writer,sheet_name='Sheet1')
writer.save()