import csv
import glob
import os
import pandas as pd
from datetime import date,datetime  
dfflag=0
absolute_path = os.path.dirname(os.path.abspath(__file__))
main = glob.glob('Monthwise_Data_Dump/*',recursive='true') # assume the path
for folder in main:
    folder2=folder+"/*.csv"
    filename=folder[(folder.find('\\'))+1:]
    output = absolute_path+"\\Monthly_Report"+'\\'+filename+".csv"
    print(output)
    for csvmonth in glob.glob(folder2):
        csvname=csvmonth[(csvmonth.rfind('\\'))+1:csvmonth.rfind('.')]
        print(csvname)
        scsvname=str(csvname)
        dfcurrfile=pd.read_csv(csvmonth)
        if(dfflag==0):
            dfcheck=pd.DataFrame(dfcurrfile)
            dfflag=1                
            continue
        dfcheck[[scsvname]]=''
        #checking unique combination of name society building flat category
        merged =dfcheck[['Name','Society','Building','Flat','Category']].astype(str).merge(dfcurrfile.astype(str),on=['Name','Society','Building','Flat','Category'],indicator=True, how='outer')
        newentry=merged.loc[merged['_merge'] == 'right_only']
        sameentry=merged.loc[merged['_merge'] != 'right_only']
        dfcheck.reset_index()
        dfcheck[['Name','Society','Building','Flat','Category',scsvname]]=sameentry[['Name','Society','Building','Flat','Category',scsvname]]
        if(newentry.empty==False):
            dfcheck=dfcheck.append(newentry[['Name','Society','Building','Flat','Category',scsvname]],ignore_index=True, sort=False)
            #dfcheck=dfcheck.reset_index().append(newentry[['Name','Society','Building','Flat','Category',scsvname]],ignore_index=True, sort=False)
        print("Done")
        print(dfcheck.to_markdown)
    dfcheck.to_csv(output,index=False)
