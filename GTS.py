import win32com.client as win32
import pandas as pd

def C08(CN,BTM):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm')
    ws = wb.Worksheets('Sheet1')

    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C08"
    ws.Range('E59').Value = 'Yes'
    ws.Range('E60').Value = 'C08 - DEA Licence/Narcotic'
    ws.Range('E61').Value = BTM
    ws.Range('E62').Value = '30.12.9999'
    ws.Range('E63').Value = 'NA'
    ws.Range('E64').Value = '86838397, 86838400, 86838419, CA576610HGE, CA781773RGE, CA781810HGE, CA7818Y10GE, GTS-BUPRENORPHONEHYDROCHLORIDE, GTS-METHADONHYDROCHLORID'
    ws.Range('E65').Value = 'GTS-BUPRENORPHONEHYDROCHLORIDE: 9,999,999 MG; GTS-METHADONHYDROCHLORID: 9,999,999 MG'

    new_file = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN}_create C08 licence.xlsm'
    wb.SaveAs(new_file)
    wb.Close(SaveChanges=False)



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

    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C06"
    ws.Range('E59').Value = 'Yes'
    ws.Range('E60').Value = 'C06 - Veterinary'
    ws.Range('E61').Value = 'NA'
    ws.Range('E62').Value = '30.12.9999'
    ws.Range('E63').Value = 'L01, L02, L03, L04, NONE'
    ws.Range('E64').Value = 'NA'
    ws.Range('E64').Value = '9,999,999'

    path = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C06 licence.xlsm'
    wb.SaveAs(path)
    wb.Close(SaveChanges=False)
    excel.Quit()



def C34(CN):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm')
    ws = wb.Worksheets('Sheet1')

    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C34"
    ws.Range('E59').Value = 'Yes'
    ws.Range('E60').Value = 'C34 - Registration for Complementary Feed for Farm Animals'
    ws.Range('E61').Value = 'NA'
    ws.Range('E62').Value = '30.12.9999'
    ws.Range('E63').Value = 'L10'
    ws.Range('E64').Value = 'NA'
    ws.Range('E65').Value = '9,999,999'

    path = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C34 licence.xlsm'
    wb.SaveAs(path)
    wb.Close(SaveChanges=False)
    excel.Quit()

def C33(CN):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm')
    ws = wb.Worksheets('Sheet1')

    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C33"
    ws.Range('E59').Value = 'Yes'
    ws.Range('E60').Value = 'C33 - Vet Samples'
    ws.Range('E61').Value = 'NA'
    ws.Range('E62').Value = '30.12.9999'
    ws.Range('E63').Value = 'NA'
    ws.Range('E64').Value = 'CA5536030GQZ1, CA5537030GQZ1, CA5538030GQZ1, CA5539030GQZ1'
    ws.Range('E65').Value = '2'

    path = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C33 licence.xlsm'
    wb.SaveAs(path)
    wb.Close(SaveChanges=False)
    excel.Quit()


C33('53633')












