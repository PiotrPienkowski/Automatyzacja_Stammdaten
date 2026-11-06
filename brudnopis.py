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
    page.locator(".active-result", has_text="HCO").click()
    time.sleep(400)


