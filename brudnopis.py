#
# from playwright.sync_api import sync_playwright
#
# with sync_playwright() as p:
#     browser = p.chromium.launch(headless=False)
#     page = browser.new_page()
#     page.goto("https://thespot.elanco.com/esc?id=sc_cat_item&sys_id=9d661f191b03d1105ca7eca3604bcb3a&sysparm_category=a20cb8eedb7c60905513c3af299619d0")
#     page.pause()


# import os
#
# prath = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze'
#
# for i in os.listdir(prath):
#     if os.path.join(prath,i).endswith('.pdf'):
#         print(i)
#     else:
#         print("png")

import xlsxwriter as xl
import pandas as pd

path = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\pharmlog 50022351.xlsx'

df = pd.read_excel(path,sheet_name='Sheet1')
numer= df[df["Kundennummer"] == 50022351]
print(numer)


