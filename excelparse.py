# coding=utf-8

from openpyxl import workbook,load_workbook


def rdexcel():
    wb = load_workbook(filename='report_20181101_193155.xlsx')
    ws = wb.active
    sheets = wb.sheetnames
    sheet_first = sheets[0]
    ws = wb[sheet_first]
    rows = ws.rows
    columns = ws.columns
    print(ws['A1'].value)
    wb.create_sheet('test', 0)
    wb.save('test.xlsx')

rdexcel()
