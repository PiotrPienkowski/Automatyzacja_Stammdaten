from playwright.sync_api import sync_playwright
import time

CN = "0050727431"

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(user_data_dir="veeva_profile",headless=False)
    page = context.new_page()
    page.goto("https://elanco.veevanetwork.com/ui/",wait_until="networkidle")
    search = page.locator(".input").first
    search.click()
    search.fill(CN)
    search.press("Enter")
    time.sleep(9999)