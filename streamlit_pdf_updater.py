from importlib.metadata import files
from operator import lshift
import os
import pandas as pd
import numpy as np
import xlsxwriter
import fitz
import time

tic = time.perf_counter()

#====================================
#------Typical Items to Edit---------
pdfDir = 'C:\\Users\\bough\\OneDrive\\Desktop\\STREAMLIT\\update\Active'
prd1 = 'All Parts'
fName = 'DF_All_Products_04/25/2024'
#---write a pickle (.sav) file in addition to xlsx
wrtPkl = True
#======================================

fOut = fName + ".xlsx"
workbook = xlsxwriter.Workbook(fOut)
worksheet = workbook.add_worksheet(prd1) 
#cols = ['OBE ITEM','DESCRIPTION','CATEGORY','MATERIAL']
cols = ['Product Name',	'Product Category',	'Document Type',	'Status',	'Document Info #1',	'Document Info #2',\
        	'Filename',	'Page Info #1',	'Page Info #2',	'Page #',	'Text',	'Filepath', 'Count']
for col_num, txt in enumerate(cols):
    worksheet.write(0,col_num,txt)
#plFile = 'Parts_List_Simple.xlsx'

pdfList = []
bolList = []
dtlList = []
row = 1
ct = -1
#for ket in range(1):
for path, catDirs, files in os.walk(pdfDir):
    #for path, prodDirs, files in 
    for pdfFile in files:
        ct += 1
        
        ln1 = ''
        ln2 = ''
        if pdfFile.endswith('pdf'):
            ff = False
            pdfList.append(pdfFile)
            # print whole path of files
            pgNm = []
            pdNm = []
            pdf = fitz.open(os.path.join(path,pdfFile))
            pg = 0
            for page in pdf:
                mf = -1
                mf2 = -1
                mf3 = -1
                pg += 1
                maxFnt = 0
                i = 0
                
                results = [] # list of tuples that store the information as (text, font size, font name) 
                dict = page.get_text("dict")
                blocks = dict["blocks"]
                for block in blocks:
                    if "lines" in block.keys():
                        spans = block['lines']
                        for span in spans:
                            data = span['spans']  
                            for lines in data:
                                #if keyword in lines['text'].lower(): # only store font information of a specific keyword
                                lineStr = lines['text'].replace('®','')
                                lineStr = lineStr.replace('™','')
                                
                                if "BuildingEnvelope" in lines['text'] or lines['text'] == 'Table of Contents':
                                    be = True
                                #elif '"' in lines['text']:
                                #    be = False
                                #elif '/' in lines['text']:
                                #    be = False
                                #elif '®' in lines['text']:
                                #    be = False
                                #elif '™' in lines['text']:
                                #    be = False
                                elif pg == 1 and '®' in lines['text']:
                                    be = False
                                elif lines['text'] == '':
                                    be = False
                                else:
                                    results.append((lines['text'], lines['size'], lines['font']))                        
                                        
                                    if lines['size'] > maxFnt:
                                        maxFnt = lines['size']
                                        mf = i
                                        mf2 = -1
                                        mf3 = -1
                                        if pg == 1:
                                            ln1 = lines['text']
                                        # lines['text'] -> string, lines['size'] -> font size, lines['font'] -> font name
                                    elif lines['size'] == maxFnt:
                                        if mf2 == -1:
                                           mf2 = i
                                        elif mf3 == -1:
                                            mf3 = i
                                        if pg == 1:
                                            ln2 = lines['text'] 
                                    i += 1
                if pg > 2 or True:              
                    #for prt in df['OBE ITEM']:
                    for j in range(len(results)):
                        #if str(prt) in str(results[j]):  
                        fldr = path.split('\\')
                        #if prt in str(results[j]): 
                        #doc_list.append(fldr[-1])
                        #prod_list.append(fldr[-2])
                        #cat_list.append(fldr[-3])
                            #print('We found a part')
                        #---product name---
                        worksheet.write(row, 0, fldr[-2])
                        #---product category---
                        worksheet.write(row, 1, fldr[-3])
                        #---document type---
                        worksheet.write(row, 2, fldr[-1])
                        #---Product status---
                        worksheet.write(row, 3, fldr[-4])

                        worksheet.write(row, 4, ln1) 
                        worksheet.write(row, 5, ln2) 
                        worksheet.write(row, 6, pdfFile) 
                        worksheet.write(row, 7, results[mf][0]) 
                        if mf2 >= 0:
                            worksheet.write(row, 8, results[mf2][0]) 
                        #if mf3 >= 0:
                        #    worksheet.write(row, 9, results[mf3][0]) 
                        worksheet.write(row, 9, pg) 
                        worksheet.write(row,10,results[j][0])
                        worksheet.write(row,11,path)
                        worksheet.write(row,12,1)

                        #for k in range(len(cols)):
                        #    worksheet.write(row,col,df[cols[k]].loc[(df['OBE ITEM'] == prt)].values[0])
                        #    col += 1
                        row += 1               
                    
                #pdNm.append(results[mF][0])
                #pgNm.append(results[mF2][0])

            pdf.close()
            bolList.append(ff)
                                
workbook.close()
print('PDF reading complete')

#---Open the excel file and write the pickle file needed for streamlit---
#---Note: it would be quicker to store all the data as a dataframe
#----to begin with and then write both an xlsx and sav at the end
#----but the code was originally written to save an excel file
#----it would need to be re-written somewhat
if wrtPkl:
    df1 = pd.read_excel(fOut)
    df1['Text'].replace('',np.nan, inplace=True)
    df1 = df1.dropna(subset=['Text'])
    pklFile = fName + ".sav"
    df1.to_pickle(pklFile)

toc = time.perf_counter()

print(f"Run time {toc - tic:0.4f} seconds")