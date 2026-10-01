import win32com
import win32com.client as win32
from playwright.sync_api import sync_playwright
import time
import os

#######prompt
# na podstawie zalaczonego dokumentu powiedz mi jaka jest lokalizacja. podzieldane na podstawie dokumentu na 3 zmienne
#street1 - tu wpisz ulice i numer domu,
# city - tu wpisz miastp
# postal_code - tu wpisz kod pocztowy. wszstkie zmienne musza byc w cudzyslowiu " np city = "Berlin"

street1 = "Hindenburgstr. 31"
city = "Bad Königshofen"
postal_code = "97631"

templatka = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm'
robocze = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'
snow_de = rf'https://thespot.elanco.com/esc?id=sc_cat_item&sys_id=9d661f191b03d1105ca7eca3604bcb3a&sysparm_category=a20cb8eedb7c60905513c3af299619d0'
snow_ch = rf'https://thespot.elanco.com/esc?id=sc_cat_item&table=sc_cat_item&sys_id=666af21697260e907487fd7ef053afd3&recordUrl=com.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1&sysparm_id=666af21697260e907487fd7ef053afd3'

taskkill = os.system('taskkill /f /im excel.exe')


class snow_ticket:

    def __init__(self,CN):
        self.CN = CN
        self.p = sync_playwright().start()
        self.context = self.p.chromium.launch_persistent_context(user_data_dir="veeva_profile", headless=False)
        self.page = self.context.new_page()
        self.new_file = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN}_customer block.xlsm'


    def snow_de(self):
        self.page.goto(snow_de,wait_until="domcontentloaded")
        self.page.locator("#s2id_sp_formfield_sales_organization a").click()
        self.page.get_by_role("option", name="DE01").click()
        self.page.locator("#s2id_sp_formfield_type_of_request a").click()
        self.page.get_by_role("option", name="Change").click()
        self.page.get_by_role("textbox", name="Customer Number").fill(self.CN)
        self.page.locator("#s2id_sp_formfield_request_priority a").click()
        self.page.get_by_role("option", name="Standard - 48 hours").click()
        self.page.locator("#s2id_sp_formfield_distribution_channel a").click()
        self.page.get_by_role("option", name="10").click()
        self.page.get_by_role("textbox", name="Additional information").fill(rf'Hi Team, please change adress (see attached)')
        self.page.locator("#s2id_sp_formfield_multiple_requests a").click()
        self.page.get_by_role("option", name="No", exact=True).click()
        self.page.locator("#s2id_sp_formfield_account_group a").click()
        self.page.get_by_role("option", name="Sold-to").click()
        self.page.locator('#cmd_form_attached input[type="file"]').set_input_files(rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN}_customer block.xlsm')
        with self.page.expect_file_chooser() as cf:
            self.page.get_by_role("button", name="Choose a file").click()
        for i in os.listdir(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'):
            file_path1 = os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze',i)
            if file_path1.lower().endswith(('.pdf','.jpg', '.png')):
                cf.value.set_files(file_path1)
                break


    def snow_ch(self):
        self.page.goto(snow_ch, wait_until="domcontentloaded")
        self.page.locator("#s2id_sp_formfield_type_of_request a").click()
        self.page.get_by_role("option", name="Change").click()
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
            if file_path.lower().endswith(('.pdf','.jpg', '.png')):
                cf1.value.set_files(os.path.join(file_path))
                break

class DE(snow_ticket):

    def __init__(self, CN, reference):
        super().__init__(CN)
        self.reference = reference
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        self.wb = excel.Workbooks.Open(templatka)
        self.ws = self.wb.Worksheets('Sheet1')
        self.ws.Range('A12').Value = 'DE01'
        self.ws.Range('B12').Value = 'Change'
        self.ws.Range('C12').Value = 'Sold-to'
        self.ws.Range('E5').Value = self.CN
        self.ws.Range('E12').Value = street1
        self.ws.Range('E14').Value = city
        self.ws.Range('E17').Value = postal_code
        self.ws.Range('E23').Value = self.reference
        self.ws.Range('E59').Value = 'No'
        self.wb.SaveAs(rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN}_adress chanhge.xlsm')
        self.wb.Close()
        excel.Application.Quit()
        self.snow_de()

        time.sleep(300)

DE("50025495", "C06")



