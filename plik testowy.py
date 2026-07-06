import pandas as pd
import win32com.client as  win32

CN = '50021166'
BTM = '25242'


excel = win32.Dispatch('Excel.Application')
wb1 = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\BTM Template.xlsx')
ws1 = wb1.Worksheets('de')
wb1.RefreshAll()
excel.CalculateUntilAsyncQueriesDone()
data1 = ws1.UsedRange.Value
df = pd.DataFrame(data1)

# pierwszy wiersz -> nagłówki
df.columns = df.iloc[0]

# usuń pierwszy wiersz
df = df.iloc[1:].reset_index(drop=True)
df = df[df['Kundennummer'] == CN]
df['KLIENT'] = '342'
df['NR_BTM'] = BTM
print(df)

