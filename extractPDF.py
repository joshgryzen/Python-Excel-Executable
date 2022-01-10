import PyPDF2 as pydf
from PyPDF2 import PdfFileReader
import glob, os, shutil, xlsxwriter
import pandas as pd
import numpy as np
import extract_msg
import win32com.client

path = os.getcwd()

try:
    finalexcelsheet = pd.read_excel(path + "\Combined Form Data.xlsx")
except:
    workbook = xlsxwriter.Workbook('Combined Form Data.xlsx')
    workbook.close()
    finalexcelsheet = pd.read_excel(path + "\Combined Form Data.xlsx")

extract = path + "\Extract Complete"
if not os.path.exists(extract):
    os.makedirs(extract)

append = path + "\Append Complete"
if not os.path.exists(append):
    os.makedirs(append)

emails = path + "\Old Emails"
if not os.path.exists(emails):
    os.makedirs(emails)

#print(path)

#for file in glob.glob(path + "\*.msg"):
    #msg = extract_msg.openMsg(file)
    #outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    #msg = outlook.OpenSharedItem(file)        
    #att = msg.attachments
    #for i in att:
        #i.SaveAsFile(os.path.join(path, i.FileName))

    #filename = file[len(path)+1:]

    #print(filename)
    #print('path: ',  path + "\Old Emails\Completed - " + filename)
    #shutil.move(file, path + "\Old Emails\Completed - " + filename)

for file in glob.glob(path + "\*.msg"):
    filename = file[len(path)+1:]
    movedFile = path + "\Old Emails\Completed - " + filename
    shutil.move(file, movedFile)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(movedFile)        
    att = msg.attachments
    for i in att:
        name = os.path.join(path, i.FileName)
        dup = 0

        def checkName(name, dup):
            for file in glob.glob(path + "\*.pdf"):
                if file == name:
                    dup = dup + 1
                    name = os.path.join(path, str(dup) + " - " + i.FileName)
                    return checkName(name, dup)
            return name

        i.SaveAsFile(checkName(name, dup))

for file in glob.glob(path + "\*.pdf"):
    
    pdf = PdfFileReader(file)  
    fields = pdf.getFields()

    name = []
    val = []

    for field_name, value in fields.items():
        field_value =  value.get('/V', None)

        if field_value == "/On": field_value = "Yes"
        #if field_value == "" and "Notes" in field_name == False: field_value = "No"
        #if field_value == None and "Notes" in field_name == False: field_value = "No"

        name.append(field_name)
        val.append(field_value)    

    npArr = np.array([name, val])
    npArr = npArr.transpose()

    filename = file[:file.index(".")]

    #print(filename)

    workbook = xlsxwriter.Workbook(filename + ".xlsx")
    worksheet = workbook.add_worksheet()

    row = 0

    for col, data in enumerate(npArr):
        worksheet.write_column(row, col, data)

    workbook.close()

    filename = file[len(path)+1:]

    #print(filename)
    #print('path: ',  path + "\Extract Complete\Completed - " + filename)
    shutil.move(file, path + "\Extract Complete\Completed - " + filename)

for file in glob.glob(path + "\*.xlsx"):
    if file != (path + "\Combined Form Data.xlsx"):
        df = pd.concat(pd.read_excel(file, sheet_name =None), ignore_index=True, sort=False)

        finalexcelsheet = finalexcelsheet.append(df, ignore_index=True)
        finalexcelsheet.to_excel(path + "\Combined Form Data.xlsx", index = False)

        filename = file[len(path)+1:]

        print(filename)
        print('path: ',  path + "\Appended Complete\Appended - " + filename)
        shutil.move(file, path + "\Append Complete\Completed - " + filename)