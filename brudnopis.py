
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto("https://thespot.elanco.com/esc?id=sc_cat_item&table=sc_cat_item&sys_id=666af21697260e907487fd7ef053afd3&recordUrl=com.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1&sysparm_id=666af21697260e907487fd7ef053afd3")
    page.pause()