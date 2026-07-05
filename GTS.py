import win32com.client as win32
import pandas as pd

def C08(CN,BTM):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm')
    ws = wb.Worksheets('Sheet1')

    ws.Range('E5').Value = CN
    ws.Range('E23').Value = "C08"
    ws.Range('E59').Value = 'yes'
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
    df =df[df[8].astype(str) == str(CN)]
    df[0] = '342'
    df[1] = BTM

    new_file1 = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\pharmlog {CN}.xlsx'
    df.to_excel(new_file1, index=False)
    wb1.Close(SaveChanges=False)
    excel.Quit()


C08('50001442','252452')






