import pandas as pd
import win32com.client

CN = '50001442'
BTM = '25242'


excel = win32com.client.Dispatch('Excel.Application')
wb1 = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\BTM Template.xlsx')
ws1 = wb1.Worksheets('de')
wb1.RefreshAll()
excel.CalculateUntilAsyncQueriesDone()
data1 = ws1.UsedRange.Value
df = pd.DataFrame(data1[2:], columns = data1[0])
df8 = df[df['Kundennummer'].astype(str) == str(CN)]
# df['KLIENT'] = '342'
# df['NR_BTM'] = BTM
print(df)