import openpyxl
from openpyxl import load_workbook
import csv
import glob
import os
import pandas as pd
from datetime import date,datetime  

absolute_path = os.path.dirname(os.path.abspath(__file__))
datefiles=glob.glob(absolute_path+'/*/')
for datefolder in datefiles:

    file1=datefolder.rfind('\\')
    #print(file1)
    file2=datefolder[:datefolder.rfind('\\')].rfind('\\')
    #print(file2)
    mainname=datefolder[file2+1:file1]
    
    if(mainname=='src'):
        excel_files = glob.glob(datefolder+'*.xlsx') # assume the path
        pathr='/src2/'
        try:
            os.makedirs(absolute_path+pathr)
        except:
            pass 
        for excel in excel_files:
            #print(excel)
            #print(excel)
            # create an excel file object
            reader = pd.ExcelFile(excel)
            wb = load_workbook(excel, data_only = True)
            # loop through all sheet names
            file1=excel.rfind('.')
            file2=excel.rfind('\\')
            name=excel[file2+1:file1]
            #print(excel)
            for sheet in reader.sheet_names:
                file1=excel.rfind('.')
                file2=excel.rfind('\\')
                name2=excel[file2+1:file1]+'_'+sheet
                fname=absolute_path +pathr+name+'.xlsx'
                #print(fname)
                #print("name:",name2)
                sh = wb[sheet]
                for i in range(ord('A'),ord('Z')):
                    for j in range(1,150):
                        y=chr(i)+str(j)
                        color_in_hex = (sh[y].fill.start_color.index) # this gives you Hexadecimal value of the color
                        if(color_in_hex!='00000000'):
                            print (y,'HEX =',color_in_hex)
                            print(sh[y].value)
                            print(sh[y].fill.start_color.tint)
                        if(color_in_hex==2 or (color_in_hex==0 and (sh[y].fill.start_color.tint)!=0.0)):
                            #if(str(sh[y]).find('^')!=-1):
                            print("--------------------")
                            print("chnge")
                            print (y,'HEX =',color_in_hex)
                            print(sh[y].value)
                            sh[y]=str(sh[y].value)+'^'
                            #print(sh[y])
                wb.save(fname)
