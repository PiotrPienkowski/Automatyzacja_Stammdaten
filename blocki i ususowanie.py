
import win32com.client
import time
from playwright.sync_api import sync_playwright
import os

path = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(de close and open) CMD_template4.1.4.xlsm'
new_file = r"C:\Users\02703821\OneDrive - Elanco\Desktop\robocze"
snow = "https://thespot.elanco.com/esc?id=sc_cat_item&sys_id=9d661f191b03d1105ca7eca3604bcb3a&sysparm_category=a20cb8eedb7c60905513c3af299619d0"
snow_ch = "https://thespot.elanco.com/esc?id=sc_cat_item&table=sc_cat_item&sys_id=666af21697260e907487fd7ef053afd3&recordUrl=com.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1&sysparm_id=666af21697260e907487fd7ef053afd3"

os.system("taskkill /F /IM excel.exe")

class DE:

    def __init__(self, CN):
        self.CN = CN
        self.excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
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
        self.wb.SaveCopyAs(rf'{new_file}\{self.CN}.xlsm')
        self.wb.Close()
        self.excel.Quit()



    def przegladarka_block_unblock(self):
        self.pw = sync_playwright().start()
        self.browser = self.pw.chromium.launch(channel="chrome",headless=False)
        self.page = self.browser.new_page()
        self.page.goto(snow)
        self.page.wait_for_load_state("networkidle")
        self.page.locator("#s2id_sp_formfield_sales_organization a").click()
        self.page.get_by_role("option", name="DE01").click()
        self.page.locator("#s2id_sp_formfield_type_of_request a").click()
        self.page.get_by_role("option", name="Block/unblock").click()
        self.page.get_by_role("textbox", name="Customer Number").fill(self.CN)
        self.page.locator("#s2id_sp_formfield_request_priority a").click()
        self.page.get_by_role("option", name="Standard - 48 hours").click()
        self.page.locator("#s2id_sp_formfield_distribution_channel a").click()
        self.page.get_by_role("option", name="10").click()
        self.page.locator("#s2id_sp_formfield_multiple_requests a").click()
        self.page.get_by_role("option", name="No", exact=True).click()
        self.page.locator("#s2id_sp_formfield_account_group a").click()
        self.page.get_by_role("option", name="Sold-to").click()
        self.page.locator('#cmd_form_attached input[type="file"]').set_input_files(rf"{new_file}\{self.CN}.xlsm")
        # self.page.locator("textbox", name="Additional information").fill(rf'Hello Team, /n Could you please for proceed wit a request?')
        time.sleep(300)

    def set_central_order_block(self):
        self.ws.Range('E13').Value = 'Set'
        self.ws.Range('E14').Value = '01 - Overall block'
        self.zapisywanie()
        self.przegladarka_block_unblock()

    def remove_central_order_block(self):
        self.ws.Range('E13').Value = 'Reset (Remove)'
        self.ws.Range('E14').Value = '01 - Overall block'
        self.zapisywanie()
        self.przegladarka_block_unblock()

    def remove_deletion_flag(self):
        self.ws.Range('E12').Value = 'Reset (Remove)'
        self.zapisywanie()
        self.przegladarka_block_unblock()

    def remove_central_order_and_deletio_flag(self):
        self.ws.Range('E12').Value = 'Reset (Remove)'

class CH(DE):

    def __init__(self,CN):
        super().__init__(CN)
        self.ws.Range('A12').Value = 'CH01'
        self.ws.Range('B12').Value = 'Block/Unblock/Deletion Flag'
        self.ws.Range('C12').Value = 'Sold-to'
        self.ws.Range('E5').Value = CN

    def przegladarka_block_unblock(self):
        self.pw = sync_playwright().start()
        self.browser = self.pw.chromium.launch(channel="chrome",headless=False)
        self.page = self.browser.new_page()
        self.page.goto(snow_ch)
        self.page.locator("#s2id_sp_formfield_type_of_request a").click()
        self.page.get_by_role("option", name="Block/Unblock").click()
        self.page.locator("#s2id_sp_formfield_account_group a").click()
        self.page.get_by_role("option", name="Z001 - Sold to").click()
        self.page.locator("#s2id_sp_formfield_region a").click()
        self.page.get_by_role("option", name="EMEA").click()
        self.page.locator("#s2id_sp_formfield_sales_organization_emea a").click()
        self.page.get_by_role("option", name="CH01").click()
        self.page.locator("#s2id_sp_formfield_request_priority a").click()
        self.page.get_by_role("option", name="Standard - 48 hours").click()
        self.page.locator("#s2id_sp_formfield_multiple_request a").click()
        self.page.get_by_role("option", name="No").click()
        ## tu powinien byc domestic
        self.page.locator('#cmd_form_attached input[type="file"]').set_input_files(rf"{new_file}\{self.CN}.xlsm")
        time.sleep(300)

DE('CH5').remove_deletion_flag()

