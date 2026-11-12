import win32com.client
import pandas as pd
import os
from playwright.sync_api import sync_playwright

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
# zmienne musz byc wypisane jedna pod druga!!! Koniecznie!!!

# Nie wypelniaj nastepujacych zmiennych.(tu zawse wpisz "") - id, Region, DMR_Reference_field_starts_with, Sales_Rep,
# Create_GTS_with_new_account,License_Type,E-invoicing, Email_Address_Notes. W polu search_terem2 wpisuj "AWRZ/TP/".

# W polu Phone_Number oraz Email_Address wpisz dane jesli wynika to z tresi maila lub dane znajduja sie w stopce.

# Tam gdzie nie musisz nic wypeniac zmiennych to ich nie wypelniaj i wstaw "" nazwy zmiennych zawsze wpisuj zawsze w "".

# Nazw zmiennych nie podawaj w cudzyslowiach podawaj zawsze w cudzyslowiach wartosci zmiennych. Zmienne umiejszczaj zawsze jedna zmienna
# pod druga zeby kazda zmienna byla w oddzielnym wierszu a zmienne musza byc zawsze wyrownane do prawej - to wazne!
# Zmienna "Name4" powinna zawsze laczyc zmienne  Name1 +" "+ Name2  +" "+ Name3!. zmienne wypisz jedna pod druga bez spacji.

# ZMIENNE ZAWSZE MUSZA BYC NAPISANE ZAWSZE JEDNA POD DRUGA- TO WAŻNE!
# nie uuzywaj czcionki proporcjonalnej i nie sklejaj skleja formatowania.
#Wynik zwróć wyłącznie jako czysty tekst. Nie używaj tabel, markdown, bloków kodu, wypunktowań ani formatowania.
#Każda zmienna musi rozpoczynać się od nowej linii.Pomiędzy zmiennymi nie może być pustych linii.Format każdej linii:
#NazwaZmiennej="wartość"
#Przykład:
#Name1="Przykładowa praktyka"
#Name2="Jan Kowalski"
#Name3=""
#Po znaku końca linii natychmiast rozpoczynaj następną zmienną.
#Nie dodawaj żadnych komentarzy, wyjaśnień, nagłówków, źródeł ani tekstu przed pierwszą zmienną i po ostatniej zmiennej.
#Zwróć wyłącznie listę zmiennych.
#WYNIK ZWRÓĆ WEWNĄTRZ BLOKU KODU TXT. KAŻDA ZMIENNA MUSI ZNAJDOWAĆ SIĘ W OSOBNYM WIERSZU. POMIĘDZY ZMIENNYMI NIE MOŻE BYĆ PUSTYCH WIERSZY. NIE WOLNO ŁĄCZYĆ ZMIENNYCH W JEDNĄ LINIĘ.

################## tu wstaw dane z prompta

id=""
Name1="DVM Ina Wiebelitz"
Name2=""
Name3=""
Name4="DVM Ina Wiebelitz"
Street_1="Kirchsteg 4"
City="Leitzkau"
Region=""
Postal_Code="39279"
search_term2="AWRZ/TP/"
Phone_Number=""
DMR_Reference_field_starts_with=""
E_Invoicing=""
Email_Address="ina.wiebelitz@web.de"
Email_Address_Notes=""
Sales_Rep=""
Create_GTS_with_new_account=""
License_Type=""
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

templatka = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm'
lista_salesow = rf'C:\Users\02703821\Elanco\DACH_SFE Team_RedData - PLZ Liste DE,AT,CH\2026 (ab 01.08.206)\2026 PLZ Übersicht D-A-CH (Stand 01.10.2026).xlsx'
robocze =rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\\'
snow_de = r"https://thespot.elanco.com/esc?id=sc_cat_item&sys_id=9d661f191b03d1105ca7eca3604bcb3a&sysparm_category=a20cb8eedb7c60905513c3af299619d0"

os.system("taskkill /F /IM EXCEL.EXE >nul 2>nul")

class Veeva:

    def __init__(self):

        self.p = sync_playwright().start()
        self.context = self.p.chromium.launch_persistent_context(user_data_dir="veeva_profile", headless=False,no_viewport= True, channel="chrome")
        self.page = self.context.new_page()
        self.page.goto("https://elanco.veevanetwork.com", wait_until="load")
        self.page.get_by_role("button", name="Add Record").click()
        self.page.locator("#JQUID1").click()
        self.page.locator("a").filter(has_text="HCP").click()
        self.search = self.page.locator("input").first
        self.page.keyboard.type("HCO")
        self.page.keyboard.press("Enter")
        self.page.locator("a").filter(has_text="Austria").click()
        self.search2 = self.page.locator("input").first
        self.page.keyboard.type("Germany")
        self.page.keyboard.press("Enter")
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
        self.id = id
        self.numer_klienta = input("please provide Customer Number: ")
        self.REGION = input("please provide Region: ")

class DE:

    def __init__(self):

        self.excel = win32com.client.Dispatch('Excel.Application')
        self.excel.Visible = True
        self.wb = self.excel.Workbooks.Open(templatka)
        self.ws = self.wb.Worksheets('Sheet1')
        self.ws.Range('A12').Value = 'DE01'
        self.ws.Range('B12').Value = 'Change'
        self.ws.Range('C12').Value = 'Sold-to'
        self.veeva = veeva


    def sales_odpowiedzialny(self):

        self.df1 = pd.read_excel(lista_salesow, sheet_name='PLZ DE')
        self.df1.columns = self.df1.iloc[0]
        self.df1 = self.df1.iloc[1:]
        self.result = self.df1[self.df1.iloc[:, 0].astype(str).str.replace(".0", "", regex=False)== str(Postal_Code)]
        return  self.result.iloc[0, 4]

    def nowe_sold_to_templatka_DE_Create(self):


        self.ws.Range('E5').Value = self.veeva.numer_klienta
        self.ws.Range('E8').Value = Name1
        self.ws.Range('E9').Value = Name2
        self.ws.Range('E10').Value = Name3
        self.ws.Range('E12').Value = Street_1
        self.ws.Range('E14').Value = City
        self.ws.Range('E16').Value = self.veeva.REGION
        self.ws.Range('E17').Value = Postal_Code
        self.ws.Range('E19').Value = search_term2
        self.ws.Range('E20').Value = Phone_Number
        self.ws.Range('E23').Value = License_Type
        self.ws.Range('E25').Value = E_Invoicing
        self.ws.Range('E26').Value = Email_Address
        self.ws.Range('E27').Value = Email_Address_Notes
        self.ws.Range('E56').Value = self.sales_odpowiedzialny()
        self.ws.Range("E59").Value = "Yes"
        self.ws.Range('E60').Value = license_types.get(License_Type, "")
        self.new_file = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{Name1}.xlsm'
        self.wb.SaveAs(self.new_file)
        self.wb.Close()
        self.excel.Quit()

class Snow:

    def __init__(self, veeva, de):

        self.veeva = veeva
        self.de = de
        self.page = self.veeva.context.new_page()

    def open_ticket(self):
        self.page.goto(snow_de, wait_until = "load")
        self.page.keyboard.press("Enter")
        self.page.locator("#s2id_sp_formfield_sales_organization a").click()
        self.page.get_by_role("option", name="DE01").click()
        self.page.locator("#s2id_sp_formfield_type_of_request a").click()
        self.page.get_by_role("option", name="Create").click()
        self.page.get_by_role("textbox", name="Customer Number").fill(veeva.numer_klienta)
        self.page.locator("#s2id_sp_formfield_request_priority a").click()
        self.page.get_by_role("option", name="Standard - 48 hours").click()
        self.page.locator("#s2id_sp_formfield_distribution_channel a").click()
        self.page.get_by_role("option", name="10-Domestic").click()
        self.page.locator("#s2id_sp_formfield_multiple_requests a").click()
        self.page.get_by_role("option", name="No", exact=True).click()
        self.page.locator("#s2id_sp_formfield_account_group a").click()
        self.page.get_by_role("option", name="Sold-to").click()
        self.page.get_by_role("button", name="Upload Attachment for CMD").click()
        with self.page.expect_file_chooser() as fc:
            self.page.get_by_role("button",name="Upload Attachment for CMD").click()
        for i in os.listdir(rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'):
            if i.endswith('xlsm'):
                fc.value.set_files(os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze',i))
        with self.page.expect_file_chooser() as cf1:
            self.page.get_by_role("button", name="Upload Attachment for VET").click()
        for i in os.listdir(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'):
            if os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze', i).endswith(('.pdf', '.jpg', '.png')):
                cf1.value.set_files(os.path.join(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze', i))
        self.page.get_by_role("textbox", name="Additional information").fill("Hi Team, /n please create new sold-to (see attached)")
        input("press enter to continue...")





veeva = Veeva()
de = DE()
de.nowe_sold_to_templatka_DE_Create()
snow = Snow(veeva, de)
snow.open_ticket()