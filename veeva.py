from playwright.sync_api import sync_playwright
import time

CN = "0050727431"

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(user_data_dir="veeva_profile",headless=False)
    page = context.new_page()
    page.goto("https://thespot.elanco.com/esc?id=sc_cat_item&sys_id=9d661f191b03d1105ca7eca3604bcb3a&sysparm_category=a20cb8eedb7c60905513c3af299619d0",wait_until="networkidle")
    page.locator("#s2id_sp_formfield_sales_organization a").click()
    page.get_by_role("option", name="DE01").click()
    time.sleep(300)
