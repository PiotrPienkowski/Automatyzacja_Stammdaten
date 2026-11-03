import win32com.client as win32
import re
from playwright.sync_api import sync_playwright
import time

report = r'https://elancoconnect.lightning.force.com/lightning/r/Report/00ORP000001DhWP/view'

p = sync_playwright().start()
context = p.chromium.launch_persistent_context(user_data_dir="snow_profile",headless=False, channel= "chrome", no_viewport=True)
page = context.new_page()
page.goto(report)
page.frame_locator("iframe").get_by_role("button",name="More Actions - Edit").click()
page.locator("iframe").content_frame.get_by_role("menuitem", name="Export").click()
page.locator("span").filter(has_text="Details OnlyExport only the").first.click()
page.get_by_text("FormatExcel Format .xlsExcel").click()
page.get_by_role("button", name="Export").click()


outlook = win32.Dispatch("Outlook.Application")
mamespace = outlook.GetNamespace("MAPI")
skrzynka = mamespace.Folders['PIOTR.PIENKOWSKI@elancoah.com']
folder = skrzynka.Folders['Skrzynka odbiorcza']
for i in folder.Items:
    if i.Subject == '[EXTERNAL]Verify your identity in Salesforce':
        break
kod = re.search(r"Verification\s*Code:\s*(\d+)", i.Body, re.IGNORECASE)

print(kod.group(1))

pole = page.frame_locator("iframe").get_by_role(
    "textbox",
    name="Verification Code"
)

pole.click()

page.keyboard.type(kod.group(1))





