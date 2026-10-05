import win32com.client
import pandas as pd
import os
from playwright.sync_api import sync_playwright
import time


# prompt
# Przeanalizuj dokukent zalaczoy w mailu (licencje) i maila i podaj mi nazwe praktyki z podzialem na zmienne Name1, Name2 oraz
# Name3. Kazdy wiersz powinien miec max 34 znaki. Jesli nie bedzie potrzeby wpelniania zmiennej wpisz przy zmiennej "". Pamietaj zeby w jednej
# ze zmiennych umiecic osobe odpowiedzialna merytorycznie (weterynarza) za ta apteke - ta informacj musi wynikac z dokumentu.
# ponadto nazwe i numer ulicy przypisz do zmiennej Street_1 , miasto gdzie znajdujesie praktyka do City, kod pocztowy do
# Postal_Code tu podaje ci liste wszystkch zmiennych
# id =
# Name1=
# Name2=
# Name3=
# Street_1=
# City=
# Region =
# Postal_Code=
# search_term2 =
# Phone_Number =
# DMR_Reference_field_starts_with =
# E_Invoicing  =
# Email_Address =
# Email_Address_Notes =
# Sales_Rep =
# Create_GTS_with_new_account =
# License_Type =

# Nie wypelniaj nastepujacych zmiennych.(tu zawse wpisz "") - id, Region, DMR_Reference_field_starts_with, Sales_Rep,
# Create_GTS_with_new_account,License_Type,E-invoicing, Email_Address_Notes. W polu search_terem2 wpisuj "AWRZ/TP/".

# W polu Phone_Number oraz Email_Address wpisz dane jesli wynika to z tresi maila lub dane znajduja sie w stopce.

# Tam gdzie nie musisz nic wypeniac zmiennych to ich nie wypelniaj i wstaw "" nazwy zmiennych zawsze wpisuj zawsze w "".

# Nazw zmiennych nie podawaj w cudzyslowiach podawaj zawsze w cudzyslowiach wartosci zmiennych. Zmienne umiejszczaj zawsze jedna zmienna
# pod druga zeby kazda zmienna byla w oddzielnym wierszu a zmienne musza byc zawsze wyrownane do prawej - to wazne!
# Zmienna "Name4" powinna zawsze laczyc zmienne  Name1 +" "+ Name2  +" "+ Name3!. zmienne wypisz jedna pod druga bez spacji.

################## tu wstaw dane z prompta

id = "45676"
Name1 = "MEDIVET"
Name2 = "Laura Elisa Hartmann"
Name3 = ""
Name4 = "MEDIVET Laura Elisa Hartmann"
Street_1 = "Asternstraße 2"
City = "Teltow"
Region = "test"
Postal_Code = "14513"
search_term2 = "AWRZ/TP/"
Phone_Number = "+49 163 2880723"
DMR_Reference_field_starts_with = ""
E_Invoicing = ""
Email_Address = "sabine.hoffmann@medivetgroup.com"
Email_Address_Notes = ""
Sales_Rep = ""
Create_GTS_with_new_account = ""
License_Type = ""

########################################################################


license_types = {
"C02": "C02- Wholesaler trade permit",
"C05": "C05- Pharmacy",
"C06": "C06-Veterinary",
"C08": "C08- DEA Licen/Narcotic",
"C09": "C09- Mail Order Pharmacy",
"C11": "C11- Wholesaler trade permit AH",
"C13": "C13- Governmental Organization",
"C14": "C14- Intercompany",
"C16": "C16- Non pharmaceuticals",
"C29": "C29- Bioide",
"C30": "C30- Manufacturer permit",
"C31": "C31- Pet Retail",
"C33": "C33- Vet Samples",
"C34": "C34- Registration for Complementary Feed for Farm Animals",
"NLR": "NLR- No licence Required",
"RX": "RX-One time prescription"
}

templatka = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(sold-to  change  DE01)   CMD_template4.1.4.xlsm'
lista_salesow = rf'C:\Users\02703821\Elanco\DACH_SFE Team_RedData - PLZ Liste DE,AT,CH\2026 (ab 01.08.206)\2026 PLZ Übersicht D-A-CH (Stand 01.10.2026).xlsx'
robocze =rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\\'

os.system("taskkill /F /IM EXCEL.EXE >nul 2>nul")

class Veeva:

    def __init__(self):


        self.p = sync_playwright().start()
        self.context = self.p.chromium.launch_persistent_context(user_data_dir="veeva_profile", headless=False,args=["--start-maximized"], no_viewport=True)
        self.page = self.context.new_page()
        self.page.goto("https://elanco.veevanetwork.com", wait_until="networkidle")
        self.page.get_by_role("button", name="Add Record").click()
        input("Click enter to continue...")
        self.page.get_by_role("button", name="Next").click()
        self.page.get_by_text("New Address").click()
        self.page.get_by_role("listitem").filter(has_text="Corporate Name*").get_by_test_id("input").fill(Name4)
        self.page.locator("li:nth-child(9) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(Phone_Number)
        self.page.locator("li:nth-child(8) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Email_Address)
        self.page.locator("li:nth-child(3) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.click()
        self.page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1iz0kpc-1 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.click()
        self.page.get_by_text("Physical", exact=True).click()
        self.page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(2) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1iz0kpc-1 > .sc-1us6m6n-2 > .sc-1us6m6n-0").click()
        self.page.get_by_text("Yes - Default", exact=True).click()
        if Phone_Number != "":
            self.page.get_by_text("Add Phone").click()
            self.page.locator(".sc-1us6m6n-0.iYbAyy").fill(Phone_Number)
        else:
            pass
        self.page.locator("li:nth-child(3) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(Street_1)
        self.page.locator("li:nth-child(4) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.click()
        self.page.locator("li:nth-child(4) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(City)
        self.page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(9) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").click()
        self.page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(9) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Postal_Code)
        self.page.locator("li:nth-child(18) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name1)
        self.page.locator("li:nth-child(20) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name2)
        self.page.locator("li:nth-child(22) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name3)
        self.page.locator("li:nth-child(27) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1iz0kpc-1 > .sc-1us6m6n-2 > .sc-1us6m6n-0").click()
        self.page.get_by_text(search_term2, exact=True).click()
        self.page.locator("#custom-fields > .sc-12dl7xo-0 > .sc-pfq3ln-0 > li > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(Name1)
        self.page.locator(".sc-12dl7xo-0 > .sc-pfq3ln-0 > li:nth-child(3) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name2)
        self.page.locator("#custom-fields > .sc-12dl7xo-0 > .sc-pfq3ln-0 > li:nth-child(5) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name3)
        time.sleep(300)

class DE:

    def __init__(self):
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.excel.Visible = True
        self.wb = self.excel.Workbooks.Open(templatka)
        self.ws = self.wb.Worksheets('Sheet1')

    def sales_odpowiedzialny(self):
        self.df1 = pd.read_excel(lista_salesow, sheet_name='PLZ DE')
        self.df1.columns = self.df1.iloc[0]
        self.df1 = self.df1.iloc[1:]
        self.result = self.df1[self.df1.iloc[:, 0].astype(str).str.replace(".0", "", regex=False)== str(Postal_Code)]
        return  self.result.iloc[0, 4]

    def nowe_sold_to_templatka_i_snowticket(self,lic_type = False):
        self.ws.Range("E5").Value = id
        self.ws.Range('E8').Value = Name1
        self.ws.Range('E9').Value = Name2
        self.ws.Range('E10').Value = Name3
        self.ws.Range('E12').Value = Street_1
        self.ws.Range('E14').Value = City
        self.ws.Range('E16').Value = Region
        self.ws.Range('E17').Value = Postal_Code
        self.ws.Range('E19').Value = search_term2
        self.ws.Range('E20').Value = Phone_Number
        self.ws.Range('E23').Value = lic_type
        self.ws.Range('E25').Value = E_Invoicing
        self.ws.Range('E26').Value = Email_Address
        self.ws.Range('E27').Value = Email_Address_Notes
        self.ws.Range('E56').Value = self.sales_odpowiedzialny()
        self.ws.Range("E59").Value = "Yes"
        self.ws.Range('E60').Value = license_types[lic_type]
        self.wb.SaveAs(rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{id} .xlsm')
        self.wb.Close()
        self.excel.Quit()

# Veeva()
DE().nowe_sold_to_templatka_i_snowticket("C06")

