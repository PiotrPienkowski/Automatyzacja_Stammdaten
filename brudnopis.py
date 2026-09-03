import win32com.client as win32
import os
from playwright.sync_api import sync_playwright


os.system("taskkill /F /IM excel.exe 1>nul 2>nul")

link_do_snow = "https://thespot.elanco.com/esc?id=sc_cat_item&sys_id=9d661f191b03d1105ca7eca3604bcb3a&sysparm_category=a20cb8eedb7c60905513c3af299619d0"
path = rf'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm'


class licencja:

    def __init__(self,CN):
        self.CN = CN
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        self.wb= excel.Workbooks.Open(path)
        self.ws = self.wb.Worksheets('Sheet1')
        self.ws.Range('A12').Value = 'DE01'
        self.ws.Range('B12').Value = 'Change'
        self.ws.Range('C12').Value = 'Sold-to'
        self.ws.Range('E5').Value = self.CN



    def snow_ticket(self):
        with sync_playwright() as p:
            self.context = p.chromium.launch_persistent_context(user_data_dir = "snow_profile", headless=False)
            self.page = self.context.new_page()
            self.page.goto(link_do_snow,wait_until= 'networkidle')
            self.tab = self.page
            self.tab.locator("#s2id_sp_formfield_sales_organization a").click()
            self.tab.get_by_role("option", name="DE01").click()
            self.tab.locator("#s2id_sp_formfield_type_of_request a").click()
            self.tab.get_by_role("option", name="Change").click()
            self.tab.get_by_role("textbox", name="Customer Number").fill(self.CN)
            self.tab.locator("#s2id_sp_formfield_request_priority a").click()
            self.tab.get_by_role("option", name="Critical - 4 hours").click()
            self.tab.locator("#s2id_sp_formfield_distribution_channel a").click()
            self.tab.get_by_role("option", name="10-Domestic").click()
            self.tab.locator("#s2id_sp_formfield_multiple_requests a").click()
            self.tab.get_by_role("option", name="No", exact= True).click()
            self.tab.locator("#s2id_sp_formfield_account_group a").click()
            self.tab.get_by_role("option", name="Sold-to").click()
            self.tab.get_by_role("textbox", name="Additional information").fill(f'Hello Team, Please create licence (See attached)')
            with self.tab.expect_file_chooser() as fc:
                self.tab.get_by_role("button", name = "Upload Attachment for CMD").click()
            file_chooser = fc.value
            file_chooser.set_files(self.sciezka)
            with self.tab.expect_file_chooser() as fca:
                self.tab.get_by_role("button", name = "Upload Attachment for VET").click()
            for i in os.listdir(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'):
                if i.endswith(('.pdf', '.jpg', '.png')):
                    license_chooser = fca.value
                    license_chooser.set_files(os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\\', i))
                    break
            input("Przegladarka otwarta. Enter aby zamknac...")

    def C33(self):
        self.ws.Range('E23').Value =  'C33'
        self.ws.Range('E59').Value = 'Yes'
        self.ws.Range('E60').Value = 'C33 - Vet Samples'
        self.ws.Range('E61').Value = 'NA'
        self.ws.Range('E62').Value  = '31.12.2026'
        self.ws.Range('E63').Value  = 'NA'
        self.ws.Range('E64').Value  = 'CA5536030GQZ1, CA5537030GQZ1, CA5538030GQZ1, CA5539030GQZ1'
        self.ws.Range('E65').Value  = '2'
        self.sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN} create {self.ws.Range("E23").Value} licence.xlsm'
        self.wb.SaveAs(self.sciezka)
        self.wb.Close()
        self.snow_ticket()



    def C06(self):
        self.ws.Range('E23').Value =  'C06'
        self.ws.Range('E59').Value = 'Yes'
        self.ws.Range('E60').Value = 'C06 - Veterinary'
        self.ws.Range('E61').Value = 'NA'
        self.ws.Range('E62').Value = '30.12.9999'
        self.ws.Range('E63').Value = 'L01, L02, L03, L04, NONE'
        self.ws.Range('E64').Value = 'NA'
        self.ws.Range('E65').Value = '9,999,999'
        self.sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN} create {self.ws.Range("E23").Value} licence.xlsm'
        self.wb.SaveAs(self.sciezka)
        self.wb.Close()
        self.snow_ticket()



    def C34(self):
        self.ws.Range('E23').Value =  'C34'
        self.ws.Range('E59').Value = 'Yes'
        self.ws.Range('E60').Value = 'C34 - Registration for Complementary Feed for Farm Animals'
        self.ws.Range('E61').Value = 'NA'
        self.ws.Range('E62').Value =  '30.12.9999'
        self.ws.Range('E63').Value = 'L10'
        self.ws.Range('E64').Value = 'NA'
        self.ws.Range('E65').Value = '9,999,999'
        self.sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN} create {self.ws.Range("E23").Value} licence.xlsm'
        self.wb.SaveAs(self.sciezka)
        self.wb.Close()
        self.snow_ticket()


licencja("50727431").C34()