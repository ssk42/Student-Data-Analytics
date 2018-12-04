import os, openpyxl, pprint, logging, ftplib, sys, traceback
from openpyxl import Workbook
wbRoster= openpyxl.load_workbook('Roster for Salesforce LOCAL 11.2.2018 changes.xlsx')
sheetRoster= wbRoster.get_sheet_by_name('Roster')

for rowNum in range(1,704):
    allOldUsername= sheetRoster['D'+str(rowNum)].value
    allNewUsername= sheetRoster['K'+str(rowNum)].value
    if(allOldUsername != allNewUsername):
        print(sheetRoster['L'+str(rowNum)].value+' '+allOldUsername+' '+allNewUsername)
