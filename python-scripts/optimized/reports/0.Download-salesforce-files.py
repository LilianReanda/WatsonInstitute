from playwright.sync_api import sync_playwright
from pathlib import Path

PROFILE = r"C:\Users\Emanuel\playwright-salesforce-profile"

REPORT_URL = (
    "https://watson.lightning.force.com/lightning/r/"
    "Report/00OQP00000F4uRB2AZ/view?queryScope=userFolders"
)

DOWNLOAD_DIR = Path("downloads")
DOWNLOAD_DIR.mkdir(exist_ok=True)

with sync_playwright() as p:

    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE,
        channel="chrome",
        headless=False,
        accept_downloads=True
    )

    page = context.new_page()

    page.goto(REPORT_URL, wait_until="networkidle")

    # Esperar que cargue Salesforce
    page.wait_for_timeout(5000)

    # Abrir menú de acciones
    page.locator("button[aria-haspopup='true']").last.click()

    # Export
    page.get_by_text("Export", exact=True).click()

    # Esperar modal
    page.get_by_text("Details Only", exact=True).click()

    with page.expect_download() as dl:
        page.get_by_role("button", name="Export").click()

    download = dl.value

    target = DOWNLOAD_DIR / download.suggested_filename
    download.save_as(str(target))

    print("Saved:", target)

    context.close()