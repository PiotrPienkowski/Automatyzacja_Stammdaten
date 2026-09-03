import time
import os
from playwright.sync_api import sync_playwright
import win32com.client as win32

template = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm'
snow = "https://thespot.elanco.com/esc?id=sc_cat_item&sys_id=9d661f191b03d1105ca7eca3604bcb3a&sysparm_category=a20cb8eedb7c60905513c3af299619d0"
snow_ch = "https://thespot.elanco.com/esc?id=sc_cat_item&table=sc_cat_item&sys_id=666af21697260e907487fd7ef053afd3&recordUrl=com.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1&sysparm_id=666af21697260e907487fd7ef053afd3"

os.system("taskkill /F /IM excel.exe >nul 2>nul")

class snow_ticket:

    def __init__(self):

        self.pw = sync_playwright().start()
        self.browser = self.pw.chromium.launch(headless=False)
        self.page = self.browser.new_page()

    def snow_de(self, request_type = "Block/unblock"):
        self.page.goto(snow, wait_until="domcontentloaded")
        self.page.get_by_role("textbox", name="Enter your email, phone, or").fill("PIOTR.PIENKOWSKI@elancoah.com")
        self.page.keyboard.press("Enter")
        self.page.locator("#s2id_sp_formfield_sales_organization a").click()
        self.page.get_by_role("option", name="DE01").click()
        self.page.locator("#s2id_sp_formfield_type_of_request a").click()
        self.page.get_by_role("option", name=request_type).click()
        self.page.get_by_role("textbox", name="Customer Number").fill(self.CN)
        self.page.locator("#s2id_sp_formfield_request_priority a").click()
        self.page.get_by_role("option", name="Standard - 48 hours").click()
        self.page.locator("#s2id_sp_formfield_distribution_channel a").click()
        self.page.get_by_role("option", name="10").click()
        self.page.get_by_role("textbox", name="Additional information").fill(rf'Hi Team, please proceed with block')
        self.page.locator("#s2id_sp_formfield_multiple_requests a").click()
        self.page.get_by_role("option", name="No", exact=True).click()
        self.page.locator("#s2id_sp_formfield_account_group a").click()
        self.page.get_by_role("option", name="Sold-to").click()
        self.page.locator('#cmd_form_attached input[type="file"]').set_input_files(self.new_file)
        with self.page.expect_file_chooser() as cf:
            self.page.get_by_role("button", name="Choose a file").click()
        for i in os.listdir(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'):
            file_path1 = os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze',i)
            if file_path1.lower().endswith('.msg'):
                cf.value.set_files(file_path1)
                break


    def snow_ch(self,request_type = "Block/unblock"):
        self.page.goto(snow_ch, wait_until="domcontentloaded")
        self.page.get_by_role("textbox", name="Enter your email, phone, or").fill("PIOTR.PIENKOWSKI@elancoah.com")
        self.page.keyboard.press("Enter")
        self.page.locator("#s2id_sp_formfield_type_of_request a").click()
        self.page.get_by_role("option", name= request_type ).click()
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
        self.page.locator('#cmd_form_attached input[type="file"]').set_input_files(self.new_file)
        self.page.get_by_role("textbox", name="Customer Number").fill(self.CN)
        self.page.get_by_role("textbox", name="Additional information").fill(rf'Hi Team, please proceed with block')
        with self.page.expect_file_chooser()as cf1:
            self.page.get_by_role("button", name="Choose a file").click()
        for i in os.listdir(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'):
            file_path = os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze',i)
            if file_path.lower().endswith('.msg'):
                cf1.value.set_files(os.path.join(file_path))
                break


        time.sleep(300)

class de(snow_ticket):

    def __init__(self, CN):
        super().__init__()
        self.CN = CN
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.Visible = True
        self.wb = self.excel.Workbooks.Open(template)
        self.ws = self.wb.Worksheets('Sheet1')
        self.CN = CN
        self.ws.Range('A12').Value = 'DE01'
        self.ws.Range('B12').Value  = 'Block/Unblock/Deletion Flag'
        self.ws.Range('C12').Value = 'Sold-to'
        self.ws.Range('E5').Value = CN

    def zapisywanie (self):
        self.new_file = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN}_customer block.xlsm'
        self.wb.SaveAs(self.new_file)
        self.wb.Close(SaveChanges=False)
        self.excel.Quit()

    def set_central_order_block(self):
        self.ws.Range('E13').Value = 'Set'
        self.ws.Range('E14').Value = '01 - Overall block'
        self.zapisywanie()
        self.snow_de()
        self.page.get_by_role("textbox", name="Additional information").fill(rf"Hi Team, for {self.CN} please set central order block")
        time.sleep(300)

    def remove_central_order_block(self):
        self.ws.Range('E13').Value = 'Reset (Remove)'
        self.ws.Range('E14').Value = '01 - Overall block'
        self.zapisywanie()
        self.snow_de()
        self.page.get_by_role("textbox", name="Additional information").fill(rf"Hi Team, for {self.CN} please remove central order block")
        time.sleep(300)

    def remove_deletion_flag(self):
        self.ws.Range('E12').Value = 'Reset (Remove)'
        self.zapisywanie()
        self.snow_de("Deletion Flag")
        self.page.get_by_role("textbox", name="Additional information").fill(rf"Hi Team, for {self.CN} please remove deletion flag")
        time.sleep(300)

    def remove_central_order_and_deletion_flag(self):
        self.ws.Range('E12').Value= 'Reset (Remove)'
        self.ws.Range('E13').Value = 'Reset (Remove)'
        self.ws.Range('E14').Value = '01 - Overall block'
        self.zapisywanie()
        self.snow_de("Deletion Flag")
        self.page.get_by_role("textbox", name="Additional information").fill(rf"Hi Team, for {self.CN} please remove central order block and deletion flag")
        time.sleep(300)

class ch(de):

    def __init__(self, CN):
        super().__init__(CN)
        self.CN = CN
        self.ws.Range('A12').Value = 'CH01'
        self.ws.Range('B12').Value = 'Block/Unblock/Deletion Flag'
        self.ws.Range('C12').Value = 'Sold-to'
        self.ws.Range('E5').Value = CN

    def set_central_order_block(self):
        self.ws.Range('E13').Value = 'Set'
        self.ws.Range('E14').Value = '01 - Overall block'
        self.zapisywanie()
        self.snow_ch()
        self.page.get_by_role("textbox", name="Additional information").fill(rf"Hi Team, for {self.CN} please set central order block")
        time.sleep(300)

    def remove_central_order_block(self):
        self.ws.Range('E13').Value = 'Reset (Remove)'
        self.ws.Range('E14').Value= '01 - Overall block'
        self.zapisywanie()
        self.snow_ch("Deletion Flag")
        self.page.get_by_role("textbox", name="Additional information").fill(rf"Hi Team, for {self.CN} please remove central order block")
        time.sleep(300)

    def remove_deletion_flag(self):
        self.ws.Range('E12').Value= 'Reset (Remove)'
        self.zapisywanie()
        self.snow_ch("Deletion Flag")
        self.page.get_by_role("textbox", name="Additional information").fill(rf"Hi Team, for {self.CN} please remove deletion flag")
        time.sleep(300)

    def remove_central_order_and_deletion_flag(self):
        self.ws.Range('E12').Value = 'Reset (Remove)'
        self.ws.Range('E13').Value = 'Reset (Remove)'
        self.ws.Range('E14').Value = '01 - Overall block'
        self.zapisywanie()
        self.snow_ch("Deletion Flag")
        self.page.get_by_role("textbox", name="Additional information").fill(rf"Hi Team, for {self.CN} please remove central order block and deletion flag")
        time.sleep(300)

ch("8").remove_central_order_and_deletion_flag()




#  metody klas de i klasy "de" i "ch"
# set_central_order_block  - d
# remove_central_order_block
# remove_deletion_flag
# remove_central_order_and_deletion_flag