import os
import glob
import csv
import xlsxwriter
from xlsxwriter.workbook import Workbook
from pandas.io.excel import ExcelWriter
import pandas as pd
import numpy as np

os.chdir ("/EagleScout")
path = '.'
extension = 'csv'
#engine = 'xlsxwriter'
engine = 'openpyxl'

workbook = Workbook('Test.xlsx')

with pd.ExcelWriter('Test.xlsx') as writer:
    df = pd.read_csv('254-18.csv')   #Read in the CSV file
    worksheet = workbook.add_worksheet('254') #This works
    df.to_excel(writer, sheet_name = ('254')) #This doesn't work
    #pd.read_csv('254-18.csv', delimiter = ",").to_excel(writer, sheet_name=('254'))#This doesn't work either
    writer.save()
workbook.close()
