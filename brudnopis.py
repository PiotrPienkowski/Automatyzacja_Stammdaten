from playwright.async_api import Keyboard
from playwright.sync_api import sync_playwright
import time

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(user_data_dir= "veeva_profile", headless=False, channel= 'chrome', no_viewport=True)
    page = context.new_page()
    page.goto("https://elanco.veevanetwork.com")
    page.get_by_role("button", name="Add Record").click()
    page.locator("#JQUID1").click()
    page.locator("a").filter(has_text="HCP").click()
    search = page.locator("input").first
    page.keyboard.type("HCO")
    page.keyboard.press("Enter")
    page.locator("a").filter(has_text="Austria").click()
    search2 = page.locator("input").first
    page.keyboard.type("Germany")
    page.keyboard.press("Enter")
    time.sleep(300)


    #
    # for i in range(inputs.count()):
    #     print(
    #         i,
    #         inputs.nth(i).evaluate("""
    #         el => ({
    #             id: el.id,
    #             name: el.name,
    #             type: el.type,
    #             placeholder: el.placeholder,
    #             value: el.value,
    #             ariaLabel: el.getAttribute('aria-label'),
    #             ariaLabelledBy: el.getAttribute('aria-labelledby')
    #         })
    #         """)
    #     )
    #
    # time.sleep(300)
