
import win32com.client
import time

path = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(de close and open) CMD_template4.1.4.xlsm'
new_file = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\.xlsm'

class DE:

    def __init__(self, CN):
        self.CN = CN
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.ScreenUpdating = False
        self.excel.DisplayAlerts = False
        self.wb = self.excel.Workbooks.Open(path)
        self.ws = self.wb.Worksheets('Sheet1')
        self.ws.Range('A12').Value = 'DE01'
        self.ws.Range('B12').Value = 'Block/Unblock/Deletion Flag'
        self.ws.Range('C12').Value = 'Sold-to'
        self.excel.Application.Run("CreatingHeader")
        self.excel.CalculateUntilAsyncQueriesDone()
        #time.sleep(1)
        self.ws.Range('E5').Value = CN


    def zapisywanie (self):
        self.wb.SaveCopyAs(rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN}.xlsm')
        self.wb.Close()
        self.excel.Quit()

    def set_central_order_block(self):
        self.ws.Range('E13').Value = 'Set'
        self.ws.Range('E14').Value = '01 - Overall block'

        self.zapisywanie()

    def remove_central_order_block(self):
        self.ws.Range('E13').Value = 'Reset (Remove)'
        self.ws.Range('E14').Value = '01 - Overall block'

        self.zapisywanie()

    def remove_deletion_flag(self):
        self.ws.Range('E12').Value = 'Reset (Remove)'

        self.zapisywanie()

class CH(DE):

    def __init__(self,CN):
        super().__init__(CN)
        self.ws.Range('A12').Value = 'CH01'
        self.ws.Range('B12').Value = 'Block/Unblock/Deletion Flag'
        self.ws.Range('C12').Value = 'Sold-to'

CH('J23').remove_deletion_flag()

