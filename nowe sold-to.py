import win32com.client
from xlsxwriter import workbook

# prompt
# przeanalizuj dokukent zalaczoy w mailu i podaj mi nazwe praktyki z podzialem na zmienne Name1, Name2 oraz Name3. kazdy powinien miec
# max 34 znaki. jesli nie bedzie potrzeby wpelniania wpisz przy zmiennej "". Pamietaj zeby w jednej ze zmiennychcch
# umiecic osobe odpowiedzialna merytorycznie (weterynarza) za ta apteke - ta informacj musi wynikac z dokumentu.
# ponadto nazwe i numer ulicy przypisz do zmiennej Street_1 , miasto gdzie znajdujesie praktyka do City, kod pocztowy do Postal_Code


path = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(sold-to  change  DE01)   CMD_template4.1.4.xlsm'

Name1="aniDOC Tierarztpraxis GmbH"
Name2="Matthias Konermann"
Name3=""
Street_1="Geringhoffstr. 52"
City="Münster"
Postal_Code="48163"



class create_Customer_de:

    def __init__(self,Name1, Name2, Name3,Street_1,City,Postal_Code):

        self.Name1 = Name1
        self.Name2 = Name2
        self.Name3 = Name3
        self.Street_1 = Street_1
        self.City = City
        self.Postal_Code = Postal_Code
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.excel.Visible = True
        self.wb = self.excel.Workbooks.Open(path)
        self.ws = self.wb.Worksheets('Sheet1')


    def wywolanie(self):
        self.ws.Range('E8').Value = self.Name1
        self.ws.Range('E9').Value = self.Name2
        self.ws.Range('E10').Value = self.Name3
        self.ws.Range('E12').Value = self.Street_1
        self.ws.Range('E14').Value = self.City
        self.ws.Range('E17').Value = self.Postal_Code

create_Customer_de(Name1,Name2,Name3,Street_1,City,Postal_Code).wywolanie()