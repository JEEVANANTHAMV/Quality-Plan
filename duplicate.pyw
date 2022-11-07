import os
from os import path
import os.path
import shutil
from win32com.client import Dispatch
import time
import openpyxl

f1 = open("inspection_qp.txt","r")
source = f1.read()
wb=openpyxl.load_workbook(source)
wk=wb.sheetnames
if "REPORT" in wk:
    #print(source)

    #print(os.path.dirname(source))
    #print((os.path.basename(source)))
    #os.chdir(os.path.dirname(source))
    #os.system('start excel.exe'+ (os.path.basename(source)))
    exit()

destination = r"C:\Solution"
shutil.copy(source, destination)
###File Copied
path2 = source

path1 = "C:\Solution"
dest = path1+"\delete.xlsx"
path1 = path1 + "\\" +os.path.basename(source)

time.sleep(3)
if path.exists(dest):
    os.remove(dest)
    
os.rename(path1,dest)
path1 = dest
xl = Dispatch("Excel.Application")
xl.Visible = False  # You can remove this line if you don't want the Excel application to be visible

wb1 = xl.Workbooks.Open(Filename=path1)
wb2 = xl.Workbooks.Open(Filename=path2)

ws1 = wb1.Worksheets["INFPR"]
ws1.Copy(Before=wb2.Worksheets(3))

ws1 = wb1.Worksheets["INFPR 11-25"]
ws1.Copy(Before=wb2.Worksheets(4))

wb2.Close(SaveChanges=True)
wb1.Close(SaveChanges=True)

wb1 = xl.Workbooks.Open(source)

wb1.Worksheets["INFPR (2)"].Name = "REPORT"
wb1.Worksheets["INFPR 11-25 (2)"].Name = "REPORT 11-25"

wb1.Close(SaveChanges=True)
xl.Quit()
time.sleep(3)
os.remove(dest)
#os.chdir(os.path.dirname(source)+source)
#os.system('start excel.exe'+os.path.dirname(source))
