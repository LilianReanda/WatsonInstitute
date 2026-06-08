from playwright.sync_api import sync_playwright

PROFILE = r"C:\Users\Emanuel\playwright-salesforce-profile"

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE,
        channel="chrome",
        headless=False
    )

    page = context.new_page()

    page.goto("https://watson.lightning.force.com")

    print("\nLog into Salesforce normally.")
    input("\nOnce you're fully logged in, press ENTER...")

    context.close()