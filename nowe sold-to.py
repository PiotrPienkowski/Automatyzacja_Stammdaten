import win32com.client
import pandas as pd
import os


# prompt
# przeanalizuj dokukent zalaczoy w mailu (licencje) i maila i podaj mi nazwe praktyki z podzialem na zmienne Name1, Name2 oraz Name3. kazdy powinien miec
# max 34 znaki. jesli nie bedzie potrzeby wpelniania wpisz przy zmiennej "". Pamietaj zeby w jednej ze zmiennychcch
# umiecic osobe odpowiedzialna merytorycznie (weterynarza) za ta apteke - ta informacj musi wynikac z dokumentu.
# ponadto nazwe i numer ulicy przypisz do zmiennej Street_1 , miasto gdzie znajdujesie praktyka do City, kod pocztowy do Postal_Code
# tu podaje ci liste wszystkch zmiennych
# Name1=
# Name2=
# Name3=
# Street_1=
# City=
# Region =
# Postal_Code=
# seatch_terem2 =
# Phone_Number =
# DMR_Reference_field_starts_with =
# E_Invoicing  =
# Email_Address =
# Email_Address_Notes =
# Sales_Rep =
# Create_GTS_with_new_account =
# License_Type =
#  tam gdzie nie musisz nich wypeniac zmiennych to ich nie wypelniaj i wstaw "" nazwy zmiennych zawsze wpisuj zawsze w ""

################## tu wstaw dane z prompta


Name1="Tierarztpraxis Dötlingen"
Name2="Dr. Thomas Möller"
Name3=""
Street_1="Am Kohlhoff 1c"
City="Dötlingen"
Region="Niedersachsen"
Postal_Code="27801"
seatch_terem2="Thomas Möller"
Phone_Number="04433-9699840"
DMR_Reference_field_starts_with=""
E_Invoicing=""
Email_Address="info@tierarztpraxis-doetlingen.de"
Email_Address_Notes=""
Sales_Rep=""
Create_GTS_with_new_account=""
License_Type="C06"


########################################################################


license_types = {
"C02": "Wholesaler trade permit",
"C05": "Pharmacy",
"C06": "Veterinary",
"C08": "DEA Licen/Narcotic",
"C09": "Mail Order Pharmacy",
"C11": "Wholesaler trade permit AH",
"C13": "Governmental Organization",
"C14": "Intercompany",
"C16": "Non pharmaceuticals",
"C29": "Bioide",
"C30": "Manufacturer permit",
"C31": "Pet Retail",
"C33": "Vet Samples",
"C34": "Registration for Complementary Feed for Farm Animals",
"NLR": "No licence Required",
"RX": "One time prescription"
}

templatka = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(sold-to  change  DE01)   CMD_template4.1.4.xlsm'
lista_salesow = rf'C:\Users\02703821\Elanco\DACH_SFE Team_RedData - 2026 (ab 01.08.206)\2026 PLZ Übersicht D-A-CH (Stand 01.08.2026).xlsx'

zamkniecie_exceli = os.system("taskkill /F /IM EXCEL.EXE >nul 2>nul")


class DE:

    def __init__(self):

        self.excel = win32com.client.Dispatch('Excel.Application')
        self.excel.Visible = True
        self.wb = self.excel.Workbooks.Open(templatka)
        self.ws = self.wb.Worksheets('Sheet1')

    def sales_odpowiedzialny(self):
        self.wb1 = pd.read_excel(lista_salesow, sheet_name='PLZ DE')
        self.df1 = pd.DataFrame(self.wb1)
        self.df1.columns = self.df1.iloc[0]
        self.df1 = self.df1.iloc[1:]
        self.result = self.df1[self.df1[0].astype(str).str.replace(".0","") == str(Postal_Code)]
        return  self.result.iloc[0, 3]

    def tworzenie(self):
        self.ws.Range('E8').Value = Name1
        self.ws.Range('E9').Value = Name2
        self.ws.Range('E10').Value = Name3
        self.ws.Range('E12').Value = Street_1
        self.ws.Range('E14').Value = City
        self.ws.Range('E16').Value = Region
        self.ws.Range('E17').Value = Postal_Code
        self.ws.Range('E19').Value = seatch_terem2
        self.ws.Range('E20').Value = Phone_Number
        self.ws.Range('E23').Value = DMR_Reference_field_starts_with
        self.ws.Range('E25').Value = E_Invoicing
        self.ws.Range('E26').Value = Email_Address
        self.ws.Range('E27').Value = Email_Address_Notes
        self.ws.Range('E56').Value = self.sales_odpowiedzialny()
        self.ws.Range("E59").Select()
        self.ws.Range('E60').Value = license_types.values("C08")

print(ws.Range("E59").Validation.Type)




DE().tworzenie()
