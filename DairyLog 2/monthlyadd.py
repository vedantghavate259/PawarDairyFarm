import csv
import glob
import os
import pandas as pd
from datetime import date, datetime


def findhead(cols, heads):
    for i in heads:
        # print("Finder",i)
        if (i[1] == cols):
            return i[0]


def writetofile(writearray):
            with open(output, 'a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(writearray)


mcowcountf = 0
mcowcounth = 0
mbuffcountf = 0
mbuffcounth = 0
mpcount = 0
mdcount = 0
mMTcount = 0
mGTcount = 0
mLNcount = 0
mKVcount = 0
mRIcount = 0
mSKcount = 0
mTKcount = 0
mKKcount = 0
mGPcount = 0
absolute_path = os.path.dirname(os.path.abspath(__file__))
path = '/Test/'
path2 = '/Monthwise_Data_Dump/'
path3 = '/Monthly_Report/'
try:
    os.makedirs(absolute_path+path)
except:
    pass
try:
    os.makedirs(absolute_path+path2)
except:
    pass
try:
    os.makedirs(absolute_path+path3)
except:
    pass
datefiles = glob.glob(absolute_path+'/*/')
print(datefiles)
for datefolder in datefiles:
    print(datefolder)
    file1 = datefolder.rfind('\\')
    print(file1)
    file2 = datefolder[:datefolder.rfind('\\')].rfind('\\')
    print(file2)
    mainname = datefolder[file2+1:file1]
    if(mainname == 'src2'):
        print("datefolder: ", datefolder)
        print("Name: ", mainname)
        excel_files = glob.glob(datefolder+'*.xlsx')  # assume the path
        print(excel_files)
        for excel in excel_files:
            # print(excel)
            # create an excel file object
            reader = pd.ExcelFile(excel)
            # loop through all sheet names
            for sheet in reader.sheet_names:
                # read in data
                file1 = excel.rfind('\\')
                file2 = excel.rfind('.')
                filename = excel[file1+1:file2]
                print("filename: ", filename)
                out = absolute_path+path+filename+'_'+sheet+'.xlsx'
                print("out: ", out)
                df = pd.read_excel(excel, sheet_name=sheet)
                # save data to excel in this location
                # '~/desktop/new files/a.xlsx', etc.
                df.to_excel(out)

        excel_files = glob.glob(absolute_path+path+'*.xlsx')  # assume the path
        for excel in excel_files:
            print("excel:",excel)
            yname = excel.rfind('.')
            out = excel[:yname]+'.csv'
            # print(out)
            df = pd.read_excel(excel)  # if only the first sheet is needed.
            df.to_csv(out)

        csv_files = glob.glob(absolute_path+path+'*.csv')  # assume the path
        print("csv_files",csv_files)
        now = datetime.now()
        # dt = now.strftime("%d_%m_%Y_%H_%M_%S")
        dt2 = now.strftime("%B_%Y")
        fmainname=now.strftime("%d.%m.%Y")
        try:
            os.makedirs(absolute_path+path2+'/'+dt2)
        except:
            pass
        # print(dt2)
        output = absolute_path+path2+dt2+'/'+fmainname+'.csv'
        print("output",output)
        line = ['Name', 'Society', 'Building', 'Flat', 'Category', fmainname]
        with open(output, 'w', newline='') as file:
            pass

        writetofile(line)

        aname = []

        for csvt in csv_files:
            print(csvt)
            file1 = csvt.rfind('\\')
            file2 = csvt.rfind('.')
            filename = csvt[file1+1:file2]
            rowc = 0
            head = []
            # print(filename)
            sheet = filename.rfind('_')
            Dboyname = filename[:sheet]
            # print(Dboyname)
            soctempname = filename[sheet+1:]
            # print(soctempname)
            socname = ''
            if(soctempname.find("Sheet") == -1):
                socname = soctempname
                # print("test",socname)
            with open(csvt, 'r') as file:
                reader = csv.reader(file)
                for row in reader:
                    rowc = rowc+1
                    colc = 0
                    for column in row:
                        cow = 0
                        buff = 0
                        cell = column
                        if(cell.find('^') != -1):
                            if(head == []):
                                head.append([cell[:cell.find('^')], colc])
                                # print("Intial")
                                # print("Head:",head)
                            else:
                                # print("Inside else")
                                loopc = 0
                                for i in head:
                                    # print(i[1])
                                    if(i[1] == colc):
                                        # print("Before",head[loopc][0])
                                        head[loopc][0] = cell[:cell.find('^')]
                                        # print("Text:",cell[:cell.find('^')])
                                        # print("After",head[loopc][0])
                                        break
                                    loopc = loopc+1
                                    # print(loopc)
                                head.append([cell[:cell.find('^')], colc])
                                # print("Append")
                                # print("Head:",head)

                        result1 = cell.find('-')
                        if(result1 != -1):
                            main = cell[result1+1:]
                            acowcount = 0.0
                            abuffcount = 0.0
                            while True:

                                result2 = main.find(' ')
                                product = main
                                if(result2 != -1):
                                    product = main[:result2]
                                item = product
                                n = len(item)
                                dec = item.find('.')
                                if(item.find('C') != -1):
                                    acowcount = acowcount+float(item[:n-1])
                                    # print(acow)
                                    # print(Dboyname,socname,findhead(colc,head),cell[:result1],"Cow",acow)
                                elif(item.find('B') != -1):
                                    abuffcount = abuffcount+float(item[:n-1])
                                    # print(Dboyname,socname,findhead(colc,head),cell[:result1],"Buffalo",item[:n-1])
                                y = main.find(' ')
                                main = main[result2+1:]
                                if(y == -1):
                                    break
                                elif(y == -1 and main.find('^') == 1):
                                    break
                            acow = [Dboyname, socname, findhead(
                                colc, head), cell[:result1], "Cow", acowcount]
                            print(acow)
                            if(acowcount != 0.0):
                                writetofile(acow)
                            abuff = [Dboyname, socname, findhead(
                                colc, head), cell[:result1], "Buffalo", abuffcount]
                            # print(abuff)
                            if(abuffcount != 0.0):
                                writetofile(abuff)
                            acowcount = 0.0
                            abuffcount = 0.0
                        colc = colc+1
                    # print("Head:",head)
        # os.startfile(output)
