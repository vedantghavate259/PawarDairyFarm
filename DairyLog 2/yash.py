import csv
import glob
import os
import pandas as pd
from datetime import date,datetime
mcowcountf =0
mcowcounth =0
mbuffcountf=0
mbuffcounth=0
mpcount=0
mdcount=0
mMTcount=0               
mGTcount=0
mLNcount=0
mKVcount=0
mRIcount=0
mSKcount=0
mTKcount=0
mKKcount=0
mGPcount=0
excel_files = glob.glob('src/*.xlsx') # assume the path
path='Test/'
for excel in excel_files:
    # create an excel file object
    reader = pd.ExcelFile(excel)
    # loop through all sheet names
    for sheet in reader.sheet_names:
        #read in data
        file1=excel.find('\\')
        file2=excel.find('.')
        filename=excel[file1+1:file2]
        out = path+filename+'_'+sheet+'.xlsx'
        df = pd.read_excel(excel,sheet_name=sheet)
        # save data to excel in this location
        # '~/desktop/new files/a.xlsx', etc.
        df.to_excel(out)
excel_files = glob.glob('Test/*.xlsx') # assume the path
for excel in excel_files:
    out = excel.split('.')[0]+'.csv'
    df = pd.read_excel(excel) # if only the first sheet is needed.
    df.to_csv(out) 
csv_files = glob.glob('Test/*.csv') # assume the path
print("\t\t\t\t\t\t\tBuffalo Full\tBuffalo Half\tCow Full\tCow Half\tDahi\tPaneer")
now = datetime.now()
dt = now.strftime("%d/%m/%Y %H:%M:%S")
absolute_path = os.path.dirname(os.path.abspath(__file__))
output = absolute_path + '/result.csv'
with open(output, 'a', newline='') as file:
    writer = csv.writer(file)
    writer.writerow([dt,'Buffalo Full','Buffalo Half','Cow Full','Cow Half','Dahi','Paneer','Buffalo Ghee',
                            'Cow Ghee','Loni','Khava','Basundi','Shrikhand','Taak','Chik','Green Peas'])
for csvt in csv_files:
    file1=csvt.find('\\')
    file2=csvt.find('.')
    filename=csvt[file1+1:file2]
    cowcountf =0
    cowcounth =0
    buffcountf=0
    buffcounth=0
    pcount=0
    dcount=0
    MTcount=0               
    GTcount=0
    LNcount=0
    KVcount=0
    RIcount=0
    SKcount=0
    TKcount=0
    KKcount=0
    GPcount=0
    #print(filename)
    sheet=filename.find('_')
    Dboyname=filename[:sheet]
    #print(Dboyname)
    soctempname=filename[sheet+1:]
    #print(soctempname)
    socname=''
    if(soctempname.find("Sheet")==-1):
        socname=soctempname
        print("test",socname)
    with open(csvt, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
                for column in row:
                    cell=column
                    result1 = cell.find('-')
                    if(result1 != -1):
                        main=cell[result1+1:]
                        while True:
                            result2=main.find(' ')
                            product = main
                            if(result2!=-1):
                                product = main[:result2]
                            item=product
                            n=len(item)
                            dec = item.find('.')
                            #if(item[n-1]=='C'):
                            if(item.find('C')!=-1):
                                if(dec==-1):
                                    fullc=int(item[:n-1])
                                    cowcountf=cowcountf+int(fullc)
                                else:
                                    fullc = int(item[:dec])
                                    cowcounth=cowcounth+1
                                    cowcountf=cowcountf+int(fullc)
                            elif(item.find('B')!=-1):
                                if(dec==-1):
                                    fullc = int(item[:n-1])
                                    buffcountf=buffcountf+int(fullc)
                                else:
                                    fullc = int(item[:dec])
                                    buffcounth=buffcounth+1
                                    buffcountf=buffcountf+int(fullc)
                            elif(item.find('D')!=-1):
                                dcount=dcount+1
                            elif(item.find('PN')!=-1):
                                pcount=pcount+1                      
                            elif(item.find('MT')!=-1):
                                MTcount=MTcount+1
                            elif(item.find('GT')!=-1):
                                GTcount=GTcount+1
                            elif(item.find('LN')!=-1):
                                LNcount=LNcount+1
                            elif(item.find('KV')!=-1):
                                KVcount=KVcount+1
                            elif(item.find('RI')!=-1):
                                RIcount=RIcount+1
                            elif(item.find('SK')!=-1):
                                SKcount=SKcount+1
                            elif(item.find('TK')!=-1):
                                TKcount=TKcount+1
                            elif(item.find('KK')!=-1):
                                KKcount=KKcount+1
                            elif(item.find('GP')!=-1): 
                                GPcount=GPcount+1
                            y = main.find(' ')
                            main=main[result2+1:]
                            if(y==-1):
                                break
        with open(output, 'a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([filename,buffcountf,buffcounth,cowcountf,cowcounth,dcount,pcount,MTcount,              
                            GTcount,LNcount,KVcount,RIcount,SKcount,TKcount,KKcount,GPcount])
        print(filename,"\t",buffcountf,"\t",buffcounth,"\t",cowcountf,"\t",cowcounth,"\t",dcount,"\t",pcount)
        mbuffcounth=mbuffcounth+buffcounth
        mbuffcountf=mbuffcountf+buffcountf
        mcowcounth=mcowcounth+cowcounth
        mcowcountf=mcowcountf+cowcountf
        mdcount=mdcount+dcount
        mpcount=mpcount+pcount
        mMTcount=mMTcount+MTcount                
        mGTcount=mGTcount+GTcount
        mLNcount=mLNcount+LNcount
        mKVcount=mKVcount+KVcount
        mRIcount=mRIcount+RIcount
        mSKcount=mSKcount+SKcount
        mTKcount=mTKcount+TKcount
        mKKcount=mKKcount+KKcount
        mGPcount=mGPcount+GPcount
with open(output, 'a', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Total",mbuffcountf,mbuffcounth,mcowcountf,mcowcounth,mdcount,mpcount,mMTcount,              
                            mGTcount,mLNcount,mKVcount,mRIcount,mSKcount,mTKcount,mKKcount,mGPcount])
    writer.writerow(["END"])
    writer.writerow("")
os.startfile(output)
os.system('python checkcolor.py')
os.system('python monthlyadd.py')
