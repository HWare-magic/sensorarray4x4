from test import *
from module1 import * 
import os
os.chdir(r'C:\Users\86136\source\repos')
#转到对应目录
import csv
import xlwt
import win32com.client as win32
txtexcel()
fname = r'C:\Users\86136\source\repos\data2.xls'
excel = win32.DispatchEx('Excel.Application')
#excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)
wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()

fname = r'C:\Users\86136\source\repos\data1.xls'
excel = win32.DispatchEx('Excel.Application')
#excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)
wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()
dataprocess()
