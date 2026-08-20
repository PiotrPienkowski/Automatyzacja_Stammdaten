#
# from playwright.sync_api import sync_playwright
# import time
#
# with sync_playwright() as p:
#     browser = p.chromium.launch(headless=False)
#     page = browser.new_page()
#     page.goto("https://login.veevanetwork.com/auth/login?retURL=https%3A%2F%2Felanco.veevanetwork.com/ui/")
#     page.pause()



# from playwright.sync_api import sync_playwright
# import time
#
# playwright = sync_playwright().start()
# browser = playwright.chromium.launch(headless=False)
# page = browser.new_page()
# page.goto("https://login.veevanetwork.com/auth/login?retURL=https%3A%2F%2Felanco.veevanetwork.com/ui/")
# page.wait_for_load_state("networkidle")
# page.get_by_text("Log in with").click()
# time.sleep(300)

# from os import path

import pandas as pd
# import os
#
# prath = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'
#
# for i in os.listdir(prath):
#     if os.path.join(prath,i).endswith('.pdf'):
#         print(i)
#     else:
#         print("png")

# import xlsxwriter as xl
# import pandas as pd
#
# path = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\pharmlog 50022351.xlsx'
#
# df = pd.read_excel(path,sheet_name='Sheet1')
# numer= df[df["Kundennummer"] == 50022351]
# print(numer)
#

# import win32com.client as win32com
# import os
# import time
# import pandas as pd
#
# os.system('taskkill /f /im excel.exe')
#
# tracker = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\GTS Bestellungen (3).xlsx'
#
# excel = win32com.Dispatch("Excel.Application")
# excel.visible= True
# wb_tracker = excel.Workbooks.Open(tracker)
# ws_tracker = wb_tracker.Sheets("Piotr- technical tab 2")
# table = ws_tracker.ListObjects("Tabela3")
# table.AutoFilter.ShowAllData()
# table.Range.AutoFilter(Field= 9, Criteria1 = "50000779")
# time.sleep(300)

# veeva = r"https://login.veevanetwork.com/auth/login?retURL=https%3A%2F%2Felanco.veevanetwork.com/ui/"
#
# playwright = sync_playwright().start()
# browser = playwright.chromium.launch(channel= "chromium", headless= False)
# page = browser.new_page()
# page.goto(veeva)
# page.get_by_role("button", name="Log in with Microsoft").click()
# time.sleep(300)


# from playwright.sync_api import sync_playwright
#
# playwright = sync_playwright().start()
# browser = playwright.chromium.launch(headless=False)
# tab = browser.new_page()
# tab.goto("https://login.veevanetwork.com/auth/login?retURL=https%3A%2F%2Felanco.veevanetwork.com/ui/")
# tab.get_by_role("button", name = "Log in with").click()


# from playwright.sync_api import sync_playwright
# import time
#
#
# playwright = sync_playwright().start()
# browser = playwright.chromium.launch(headless=False)
# page = browser.new_page()
# page.goto("https://elancoconnect.lightning.force.com/lightning/page/home")
# time.sleep(300)


from openpyxl import load_workbook
import time
import os

os.system("taskkill /F /IM excel.exe 1>nul 2>nul")

formatka = r'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\CMD_template4.1.4.xlsm'
nowy_plik = r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\test.xlsm'

def funkcja(CN):

    excel = load_workbook(formatka, keep_vba=True)
    ws = excel.worksheets[0]
    ws['A12'] = 'CH01'
    ws['B12'] = 'Block/Unblock/Deletion Flag'
    ws['C12'] = 'Sold-to'
    ws['E5'] = CN
    ws['E23']= "test"

    excel.save(nowy_plik)
    excel.close()

funkcja('moninski')




