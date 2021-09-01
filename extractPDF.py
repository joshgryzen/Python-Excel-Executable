import PyPDF2 as pydf
from PyPDF2 import PdfFileReader
import glob, os, shutil, xlsxwriter
import pandas as pd
import numpy as np

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

#print(path)

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

    print(filename)
    print('path: ',  path + "\Extract Complete\Completed - " + filename)
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