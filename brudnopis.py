import win32com.client as win32
tracker = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\GTS Bestellungen (3).xlsx'

excel1 = win32.Dispatch('Excel.Application')
excel1.Visible = True
wb_tracker = excel1.Workbooks.Open(tracker)
ws_tracker = wb_tracker.Worksheets('Piotr- technical tab 2').Activate()
input("czekaj")