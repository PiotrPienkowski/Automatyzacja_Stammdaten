from playwright.sync_api import sync_playwright
import time


# prompt
# przeanalizuj dokukent zalaczoy w mailu (licencje) i maila i podaj mi nazwe praktyki z podzialem na zmienne Name1, Name2 oraz
# Name3. kazdy powinien miec max 34 znaki. jesli nie bedzie potrzeby wpelniania wpisz przy zmiennej "". Pamietaj zeby w jednej
# ze zmiennych umiecic osobe odpowiedzialna merytorycznie (weterynarza) za ta apteke - ta informacj musi wynikac z dokumentu.
# ponadto nazwe i numer ulicy przypisz do zmiennej Street_1 , miasto gdzie znajdujesie praktyka do City, kod pocztowy do
# Postal_Code tu podaje ci liste wszystkch zmiennych
# Name1=
# Name2=
# Name3=
# Name4 =
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
# nie wypelniaj nastepujacych zmiennych.(tu zawse wpisz "") -Name4, Region, DMR_Reference_field_starts_with, Sales_Rep,
# Create_GTS_with_new_account,License_Type. w polu seatch_terem2 wpisuj "AWRZ/TP/". Jesli wynika to z treesci maila i
# kliet chce zalozyc e-invoicig wpisz "Yes - with pdf" ,w zmiennej Email Address wpisz maila ktory klient chce wpisac
# jako mail do odbioru fv elektoroncznych a w zmiennej "Email Address Notes'' wpisz "EDOC_DE".
#  tam gdzie nie musisz nich wypeniac zmiennych to ich nie wypelniaj i wstaw "" nazwy zmiennych zawsze wpisuj zawsze w "".
# nazw zmiennych nie podawaj w cudzyslowiach podawaj zawsze w cudzyslowiach wartosci zmiennych.  jesli w mailu jest numer
# telefonu do klienta wpisz go w zmiennej "Phone_Number". zmienne umiejszczaj zawsze jedna zmienna pod druga zeby
# kazda zmienna byla w oddzielnym wierszu a zmienne musza byc zawsze wyrownane do rpawej - to wazne! zmienna "Name4"
# powinna zawsze wygladac tak Name4 = Name1 +" "+ Name2  +" "+ Name3


#########################


Name1 = "Tierärztliche Hausapotheke"
Name2 = "Dr. Nina Keisers"
Name3 = ""
Name4 = Name1 +" "+ Name2 +" "+ Name3
Street_1 = "Thorenstr. 16"
City = "Alpen Veen"
Region = ""
Postal_Code = "46519"
seatch_terem2 = "AWRZ/TP/"
Phone_Number = ""
DMR_Reference_field_starts_with = ""
E_Invoicing = ""
Email_Address = ""
Email_Address_Notes = ""
Sales_Rep = ""
Create_GTS_with_new_account = ""
License_Type = ""






#################################

p = sync_playwright().start()
context = p.chromium.launch_persistent_context(user_data_dir="veeva_profile", headless=False,args=["--start-maximized"],no_viewport=True)
page = context.new_page()
page.goto("https://elanco.veevanetwork.com",wait_until="networkidle")
page.get_by_role("button", name="Add Record").click()
input("Click enter to continue...")
page.get_by_role("button", name="Next").click()
page.get_by_text("New Address").click()

# to jest header w veeva

page.get_by_role("listitem").filter(has_text="Corporate Name*").get_by_test_id("input").fill(Name4)
page.locator("li:nth-child(9) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(Phone_Number)
page.locator("li:nth-child(8) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Email_Address)
page.locator("li:nth-child(3) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.click()
page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1iz0kpc-1 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.click()
page.get_by_text("Physical", exact = True).click()
page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(2) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1iz0kpc-1 > .sc-1us6m6n-2 > .sc-1us6m6n-0").click()
page.get_by_text("Yes - Default", exact = True).click()
if Phone_Number != "":
    page.get_by_text("Add Phone").click()
    page.locator(".sc-1us6m6n-0.iYbAyy").fill(Phone_Number)
else:
    pass
page.locator("li:nth-child(3) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(Street_1)
page.locator("li:nth-child(4) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.click()
page.locator("li:nth-child(4) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(City)
page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(9) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").click()
page.locator(".sc-4xhgtf-0 > .sc-pfq3ln-0 > li:nth-child(9) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Postal_Code)
page.locator("li:nth-child(18) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name1)
page.locator("li:nth-child(20) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name2)
page.locator("li:nth-child(22) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name3)
page.locator("li:nth-child(27) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1iz0kpc-1 > .sc-1us6m6n-2 > .sc-1us6m6n-0").click()
page.get_by_text(seatch_terem2, exact=True).click()
page.locator("#custom-fields > .sc-12dl7xo-0 > .sc-pfq3ln-0 > li > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").first.fill(Name1)
page.locator(".sc-12dl7xo-0 > .sc-pfq3ln-0 > li:nth-child(3) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name2)
page.locator("#custom-fields > .sc-12dl7xo-0 > .sc-pfq3ln-0 > li:nth-child(5) > .sc-17jockg-10 > .sc-h9b1rc-2 > .sc-1us6m6n-2 > .sc-1us6m6n-0").fill(Name3)

# page.pause()
time.sleep(500)







# get_by_text("Record Type", exact=True)
# locator("a").filter(has_text="HCO")

# get_by_text("Primary Country", exact=True)
# locator("a").filter(has_text="Germany")