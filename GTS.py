import win32com.client as win32
import pandas as pd

def C08(CN,BTM):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm')
    ws = wb.Worksheets('Sheet1')

# runowanie makra vba

    ws.Range('A12').Value = 'DE01'
    ws.Range('B12').Value = 'Change'
    ws.Range('C12').Value = 'Sold-to'
    excel.Application.Run("CreatingHeader")

    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C08"
    ws.Range('E59').Value = 'Yes'
    ws.Range('E60').Value = 'C08 - DEA Licence/Narcotic'
    ws.Range('E61').Value = BTM

    new_file = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN}_create C08 licence.xlsm'
    wb.SaveAs(new_file)
    wb.Close(SaveChanges=False)


    wb1 = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\BTM Template.xlsx')
    ws1 = wb1.Worksheets(1)
    wb1.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()
    data1 = ws1.UsedRange.Value
    df = pd.DataFrame(data1)

    # pierwszy wiersz -> nagłówki
    df.columns = df.iloc[0]

    # usuń pierwszy wiersz
    df = df.iloc[1:].reset_index(drop=True)
    df = df[df['Kundennummer'].astype(str).str.replace('.0', '', regex=False) == CN]
    df['KLIENT'] = '342'
    df['NR_BTM'] = BTM

    new_file1 = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\pharmlog {CN}.xlsx'
    df.to_excel(new_file1, index=False)
    wb1.Close(SaveChanges=False)
    excel.Quit()


def C06(CN):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm')
    ws = wb.Worksheets('Sheet1')

    # runowanie makra vba

    ws.Range('A12').Value = 'DE01'
    ws.Range('B12').Value = 'Change'
    ws.Range('C12').Value = 'Sold-to'
    excel.Application.Run("CreatingHeader")

    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C06"
    ws.Range('E59').Value = 'Yes'
    ws.Range('E60').Value = 'C06 - Veterinary'

    path = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C06 licence.xlsm'
    wb.SaveAs(path)
    wb.Close(SaveChanges=False)
    excel.Quit()


def C34(CN):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm')
    ws = wb.Worksheets('Sheet1')

    ws.Range('A12').Value = 'DE01'
    ws.Range('B12').Value = 'Change'
    ws.Range('C12').Value = 'Sold-to'
    excel.Application.Run("CreatingHeader")


    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C34"
    ws.Range('E59').Value = 'Yes'
    ws.Range('E60').Value = 'C34 - Registration for Complementary Feed for Farm Animals'

    path = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C34 licence.xlsm'
    wb.SaveAs(path)
    wb.Close(SaveChanges=False)
    excel.Quit()

def C33(CN):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm')
    ws = wb.Worksheets('Sheet1')

    ws.Range('A12').Value = 'DE01'
    ws.Range('B12').Value = 'Change'
    ws.Range('C12').Value = 'Sold-to'
    excel.Application.Run("CreatingHeader")

    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C33"
    ws.Range('E59').Value = 'Yes'
    ws.Range('E60').Value = 'C33 - Vet Samples'

    path = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C33 licence.xlsm'
    wb.SaveAs(path)
    wb.Close(SaveChanges=False)
    excel.Quit()

#######################

C34('50017859')