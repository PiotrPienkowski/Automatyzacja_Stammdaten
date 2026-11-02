import win32com.client as win32
import pandas as pd
from playwright.sync_api import sync_playwright
import os


path = r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'
link_do_snow = "https://thespot.elanco.com/esc?id=sc_cat_item&sys_id=9d661f191b03d1105ca7eca3604bcb3a&sysparm_category=a20cb8eedb7c60905513c3af299619d0"
tracker_GTS = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\GTS Bestellungen (4).xlsx'
LS_monitoring = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\LS monitoring.xlsx'
veeva_website = rf"https://elanco.veevanetwork.com/ui/"
BTM_Template = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\BTM Template.xlsx'
CMD_template = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm'

result = os.system("taskkill /F /IM excel.exe 1>nul 2>nul") ## >nul (czarna dziura nic nie wyswietla) - przekierowuje standardowe komunikaty do "kosza" (1),przekierowuje komunikaty błędów do kosza 2)
if result != 0:  #0 to jest polecenie wykonane poprawnie tzn. procesy zamkniete <>0 blad systemwy ale nie blad obslugiwany przez sxcept
    print("no excel has been closed")
else:
    print("excel file has been closed")

class excel:

    def __init__(self):
        self.excel = win32.Dispatch('Excel.Application')
        self.excel.Visible = True

    def C08 (self,CN, BTM, Name1 = False, Name2 = False, Name3 = False, Street = False , PLZ = False, City = False):
        self.wb = self.excel.Workbooks.Open(CMD_template)
        self.ws = self.wb.Worksheets('Sheet1')
        self.ws.Range('A12').Value = 'DE01'
        self.ws.Range('B12').Value = 'Change'
        self.ws.Range('C12').Value = 'Sold-to'
        for attr, value in zip(["CN", "BTM", "Name1", "Name2", "Name3", "Street", "PLZ", "City"],[CN, BTM, Name1, Name2, Name3, Street, PLZ, City]):
            setattr(self, attr, value)
        for cell, value in zip(["E5", "E8", "E9", "E10", "E12", "E17", "E14"],[CN, Name1, Name2, Name3, Street, PLZ, City]):
            if value != False:
                self.ws.Range(cell).Value = value
            else:
                pass
        self.ws.Range('E23').Value = BTM
        if self.BTM != False:
            self.ws.Range('E59').Value = 'Yes'
            self.ws.Range('E60').Value = 'C08 - DEA Licence/Narcotic'
            self.ws.Range('E61').Value = BTM
            new_file = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN}.xlsm'
        else:
            self.ws.Range('E59').Value = 'No'
            new_file = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN}.xlsm'
        self.new_file = new_file
        self.wb.CheckCompatibility = False
        self.wb.SaveAs(new_file)
        self.wb.Close(SaveChanges=False)


    def pharmog(self):

        self.wb1 = self.excel.Workbooks.Open(BTM_Template)
        self.ws1 = self.wb1.Worksheets(1)
        self.wb1.RefreshAll()
        self.excel.CalculateUntilAsyncQueriesDone()
        self.data = self.ws1.UsedRange.Value
        self.df = pd.DataFrame(self.data)
        self.df.columns = self.df.iloc[0]
        self.df = self.df.iloc[2:].reset_index(drop=True)
        self.df = self.df[self.df['Kundennummer'].astype(str).str.replace('.0', '', regex=False) == self.CN]
        self.df['KLIENT'] = '342'
        self.df['NR_BTM'] = self.BTM
        self.new_file1 = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\pharmlog {self.CN} {self.BTM}.xlsx'
        self.df.to_excel(self.new_file1, index=False)
        self.wb1.Close(SaveChanges=False)

    def tracker_GTS(self):

        self.wb2 = self.excel.WorkBooksOpen(tracker_GTS)
        self.ws2 = self.wb2.Worksheets('Piotr- technical tab 2')
        self.data = self.ws2.UsedRange.Value
        self.df = pd.DataFrame(self.data)



class outlook:
    def __init__(self,excel_obj):

        self.excel_obj = excel_obj

        sync_playwright()
        outlook = win32.Dispatch('Outlook.Application')
        new_mail = outlook.CreateItem(0)
        new_mail.To = 'tl@pharmlog.de;btm@pharmlog.de'
        new_mail.CC = 'stammdaten@elancoah.com'
        new_mail.Subject = (f'{self.excel_obj.CN} BTM')
        new_mail.Body = (f'Guten Tag,\n '
                         f'Anbei BTM \n\n\n\n '
                         f"Mit freundlichen Grüßen\n\n"
                         f"Piotr Pieńkowski\n"
                         f"Master Data Manager DACH\n\n"
                         f"Elanco Deutschland GmbH\n"
                         f"Rolf-Schwarz-Schütte-Platz 2 / CC A01 1.OG\n"
                         f"D-40789 Monheim am Rhein\n"
                         f"piotr.pienkowski@elancoah.com\n"
                         f"www.elanco.de\n\n"
                         f"AG Bad Homburg HRB 13307 | Geschäftsführung:\n"
                         f"Dr. Inga Drosse und Cindy Wichmann\n"
                         f"Sitz der Gesellschaft: Bad Homburg\n\n"
                         f"CONFIDENTIALITY NOTICE: This email message (including all attachments) "
                         f"is for the sole use of the intended recipient(s) and may contain confidential "
                         f"and privileged information. Any unauthorized review, use, disclosure, "
                         f"copying or distribution is strictly prohibited. If you are not the intended "
                         f"recipient, please contact the sender by reply email and destroy all copies "
                         f"of the original message.\n\n"
                         f"PRIVACY NOTICE: Your privacy is important to us. To find out more about "
                         f"the information that Elanco may collect, how we use it, how we protect it, "
                         f"and your rights and choices with respect to your Personal Data, go to "
                         f"privacy.elanco.com"
                         )

        new_mail.Attachments.Add(self.excel_obj.new_file)

        for file in os.listdir(path):
            if file.lower().endswith(('.pdf','.jpg', '.png')):
                new_mail.Attachments.Add(os.path.join(path, file))
        new_mail.Display()

class snow:

    def __init__(self,excel_obj):

        self.excel_obj = excel_obj

        playwright = sync_playwright().start()
        context = playwright.chromium.launch_persistent_context(user_data_dir="snow_profile", headless=False)
        page = context.new_page()
        page.goto(link_do_snow, wait_until="load")
        page.keyboard.press("Enter")
        page.locator("#s2id_sp_formfield_sales_organization a").click()
        page.get_by_role("option", name="DE01").click()
        page.locator("#s2id_sp_formfield_type_of_request a").click()
        page.get_by_role("option", name="Change").click()
        page.get_by_role("textbox", name="Customer Number").fill(self.excel_obj.CN)
        page.locator("#s2id_sp_formfield_request_priority a").click()
        page.get_by_role("option", name="Critical - 4 hours").click()
        page.locator("#s2id_sp_formfield_distribution_channel a").click()
        page.get_by_role("option", name="10-Domestic").click()
        page.locator("#s2id_sp_formfield_multiple_requests a").click()
        page.get_by_role("option", name="No", exact=True).click()
        page.locator("#s2id_sp_formfield_account_group a").click()
        page.get_by_role("option", name="Sold-to").click()
        page.get_by_role("button", name = "Upload Attachment for VET").click()
        if any(i !=False for i in [excel_obj.Name1, excel_obj.Name2, excel_obj.Name3, excel_obj.Street, excel_obj.PLZ, excel_obj.City]):
            page.get_by_role("textbox", name="Additional information").fill(f'Hello Team, Please create C08 licence (See attached) and change data')
        else:
            page.get_by_role("textbox", name = "Additional information").fill(f'Hello Team, Please create C08 licence (See attached) and attach licence to list in DMR')
        with page.expect_file_chooser() as fc:
            page.get_by_role( "button",name = "Upload Attachment for CMD").click()
            fc.value.set_files(self.excel_obj.new_file)
        with page.expect_file_chooser()as cf1:
            page.get_by_role("button", name="Upload Attachment for VET").click()
        for i in os.listdir(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'):
            if os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze', i).endswith(('.pdf','.jpg', '.png')):
                cf1.value.set_files(os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze', i))
                break
        input("Nacisnij Enter...")
        context.close()
        playwright.stop()

class veeva:

    def __init__(self,excel_obj):

        self.excel_obj = excel_obj

        playwright = sync_playwright().start()
        context = playwright.chromium.launch_persistent_context(headless=False, args=["--start-maximized"], user_data_dir="veeva_profile")
        page = context.new_page()
        page.goto(veeva_website, wait_until="load")
        search = page.locator(".input").first
        search.click()
        search.fill("00"  + self.excel_obj.CN)
        search.press("Enter")
        input("press enter")



excel_obj = excel()
excel_obj.C08("50621810", "43", Name1 = "mimimomo")
excel_obj.pharmog()
snow(excel_obj)
veeva(excel_obj)