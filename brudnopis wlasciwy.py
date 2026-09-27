from playwright.sync_api import sync_playwright
import time

Name1 = "Tierarztpraxis Zauner"
Name2 = "Ann-Sophie Zauner"
Name3 = ""
Name_4 = Name1+ " "+Name2+ " "+Name3
Street_1 = "Marktplatz 11"
City = "Suhlendorf"
Region = ""
Postal_Code = "29562"
seatch_terem2 = "AWRZ/TP/"
Phone_Number = "+49 5820 383"
DMR_Reference_field_starts_with = ""
E_Invoicing = ""
Email_Address = ""
Email_Address_Notes = ""
Sales_Rep = ""
Create_GTS_with_new_account = ""
License_Type = ""




p = sync_playwright().start()
context = p.chromium.launch_persistent_context(user_data_dir="veeva_profile", headless=False)
page = context.new_page()
page.goto("https://elanco.veevanetwork.com",wait_until="networkidle")
page.get_by_role("button", name="Add Record").click()
input("Click enter to continue...")
page.get_by_role("button", name="Next").click()
page.get_by_text("New Address").click()

# page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(2) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1iz0kpc-1 > .sc-1us6m6n-2 > .sc-1us6m6n-0").click()
# page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(2) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1iz0kpc-1 > .sc-1us6m6n-2 > .sc-1us6m6n-0").select_option(label = "Yes - Default")
# page.locator("li:nth-child(3) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.click()
# page.locator("li:nth-child(3) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(Street_1)
# page.locator("li:nth-child(4) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.click()
# page.locator("li:nth-child(4) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(City)
# page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(9) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").click()
# page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(9) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Postal_Code)

page.pause()
# time.sleep(500)







# get_by_text("Record Type", exact=True)
# locator("a").filter(has_text="HCO")

# get_by_text("Primary Country", exact=True)
# locator("a").filter(has_text="Germany")