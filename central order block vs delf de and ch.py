import shutil
import pandas as pd
import datetime as dt
import win32com.client as  win32
import os

##1  Definiujemy zmienne globalne


sciezka = r'C:\Users\02703821\Elanco\CS-SD-DB Exchange - General'
link_dla_zespolow = r'https://elancoah.sharepoint.com/sites/CS-SD-DBExchange/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FCS%2DSD%2DDBExchange%2FShared%20Documents%2FGeneral&viewid=1b4638f8%2Df1c6%2D41dc%2Dae6d%2D04c3dfb7ec95&FolderCTID=0x012000C4F2DFAEE10E0942AEE5E850EF5504DE'
today = dt.date.today()

##2. usuwamy wszystkie pliki znajdujace sie w katalogu

for i in os.listdir(sciezka):
    pelna_sciezka = os.path.join(sciezka, i)

    if pelna_sciezka.endswith('.xlsx'):
        os.remove(pelna_sciezka)

##3. zaczybamy z plikiem DE

# otwiera plik znajdujacy sie na sharepoincie i majacy dostep do modelu semantycznego DE
excel = win32.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CS-SD-DB Exchange - General\source file\semantic_model_de.xlsx')

# klika przycisk refresh all
wb.RefreshAll()

# czekaj az odswiezanie sie zakonczy
excel.CalculateUntilAsyncQueriesDone()

ws = wb.Worksheets('central order block vs delf de')

# UsedRange oznacza wszystkie komórki, które zawierają dane.
# Value pobiera wartości z tych komórek.
data = ws.UsedRange.Value

df = pd.DataFrame(data)
df = df[df[0] == 'yes']

df['Customer Service'] = ""
df['Cash App'] = ""
df['Credit Team'] = ""
df = df.iloc[:, 1:]

df.to_excel(rf"C:\Users\02703821\Elanco\CS-SD-DB Exchange - General\DE {today}.xlsx", index=False)
wb.Close(SaveChanges=False)  # mowi ze nie zapisuje zmiany w zmiennej wb


## 4.zaczynamy w plikiem ch

# excel = win32.Dispatch("Excel.Application")
# excel.Visible = False
wb1 = excel.Workbooks.Open(r'C:\Users\02703821\Elanco\CS-SD-DB Exchange - General\source file\semantic model ch.xlsx')

wb1.RefreshAll()
excel.CalculateUntilAsyncQueriesDone()

ws1 = wb1.Worksheets('centralorderblockvsdelfch')

data1 = ws1.UsedRange.Value
df1 = pd.DataFrame(data1)
df1 = df1[df1[0] == 'yes']

df1['Customer Service'] = ""
df1['Cash App'] = ""
df1['Credit Team'] = ""
df1 = df1.iloc[:, 1:]
df1.to_excel(rf"C:\Users\02703821\Elanco\CS-SD-DB Exchange - General\CH {today}.xlsx", index=False)
wb1.Close(SaveChanges=False)
excel.Quit()

outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Display()
mail.To = "kundenbetreuung@elancoah.com","sebastian.lichota@elancoah.com","malgorzata.janaszewska@elancoah.com","monika.saroma@elancoah.com","pawel.nawrot@elancoah.com","dolzodmaa.yadamtsoo@elancoah.com","malgorzata.trznadel@elancoah.com","debitorenbuchhaltung@elancoah.com","credit_warsaw@elancoah.com"
mail.subject = rf"Central Order Block vs delf de and ch  {today}"
mail.body = mail.body = (f"Hello All, \n\n"
                         f"Please find central order block for DE and CH {today}. \n\n"
                         f'ink to SharePoint:\n {link_dla_zespolow} \n\n'
                         f'please mark "X and let me know when done so that I could set deletion flag \n\n\n\n\n\n'


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

mail.Display()
##