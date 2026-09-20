import win32com
import win32com.client as win32
import time


# import pandas as pd
# import time
#
# CN = "50000779"
#
# excel = win32.Dispatch('Excel.Application')
# excel.Visible = True
# wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\GTS Bestellungen (3).xlsx')
# ws = wb.Worksheets('Piotr- technical tab 2')
# time.sleep(1)
# ws.Activate()
# table = ws.ListObjects('Tabela3')
# if table.AutoFilter.FilterMode:
#     table.AutoFilter.ShowAllData()
# table.Range.AutoFilter(Field = 9, Criteria1 = CN)

import pandas as pd
lista_salesow = rf'C:\Users\02703821\Elanco\DACH_SFE Team_RedData - 2026 (ab 01.08.206)\2026 PLZ Übersicht D-A-CH (Stand 01.08.2026).xlsx'
PLZ = '16816'


def funkcja():
    wb1 = pd.read_excel(lista_salesow, sheet_name='PLZ DE')
    df1 = pd.DataFrame(wb1)
    df1.columns = df1.iloc[0]
    df1 = df1.iloc[1:]
    result = df1[df1['PLZ'].astype(str).str.replace(".0", "") == str(PLZ)]
    return result.iloc[0, 3]

print(funkcja())

#
# excel = win32com.client.Dispatch("Excel.Application")
# excel.Visible = True
# wb = excel.Workbooks.Open(path)
# ws = wb.Worksheets[0]
# data = ws.UsedRange.Value
# df = pd.DataFrame(data)
# df.columns = df.iloc[0]
# df = df.iloc[1:]
# df['PLZ'] = df['PLZ'].astype(str).str.replace(".0","",regex=False)
# result = df[df['PLZ'] == PLZ]
# wartosc = result.iloc[0,3]
# pd.set_option('display.max_columns', None)
# print(wartosc)
#
