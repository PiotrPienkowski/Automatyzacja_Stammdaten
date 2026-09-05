import win32com.client as win32
import pandas as pd
import time

CN = "50000779"

excel = win32.Dispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\GTS Bestellungen (3).xlsx')
ws = wb.Worksheets('Piotr- technical tab 2')
time.sleep(1)
ws.Activate()
table = ws.ListObjects('Tabela3')
if table.AutoFilter.FilterMode:
    table.AutoFilter.ShowAllData()
table.Range.AutoFilter(Field = 9, Criteria1 = CN)