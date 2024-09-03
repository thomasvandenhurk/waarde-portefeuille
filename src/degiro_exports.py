from playwright.sync_api import sync_playwright
import pyotp
import os
from datetime import datetime as dt


def login_degiro(p):
    """
    Login to DEGIRO page.
    """

    browser = p.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    page.goto("https://trader.degiro.nl/login")

    try:
        page.click('button:has-text("Alles accepteren")')
    except:
        print("No cookies button found or clicked automatically by the website.")

    # Fill in login form
    page.fill('input[name="username"]', os.environ.get('degiro-login'))
    page.fill('input[name="password"]', os.environ.get('degiro-password'))

    # Click the login button
    page.click('button[type="submit"]')

    #2FA
    totp = pyotp.TOTP(os.environ.get('totp-secret'))
    totp_code = totp.now()
    page.wait_for_selector('input[name="oneTimePassword"]')
    page.fill('input[name="oneTimePassword"]', totp_code)
    page.click('button[type="submit"]')

    return page, context


def get_session_id(page, context):
    """
    Get session id to make exports later.
    """

    page.goto('https://trader.degiro.nl/trader/#/portfolio/assets')
    page.wait_for_timeout(500)
    cookies = context.cookies()

    # Find session ID in cookies
    session_id = None
    for cookie in cookies:
        if cookie['name'] == 'JSESSIONID':
            session_id = cookie['value']
            break

    return session_id


def make_export(page, export_link: str, output_path: str):
    """
    Make export based on export link and destination folder.
    """

    with page.expect_download() as download_info:
        try:
            page.goto(export_link)
        except:
            download = download_info.value
            # Save the downloaded file
            download.save_as(output_path)
    page.wait_for_timeout(6000)


def update_exports_degiro():
    """
    Update export folder DEGIRO. This will update the (current) portfolio and rekeningoverzicht (Account)
    """

    with sync_playwright() as p:
        page, context = login_degiro(p)
        page.wait_for_timeout(500)
        session_id = get_session_id(page, context)
        page.wait_for_timeout(500)
        date = dt.now().date()

        make_export(
            page=page,
            export_link=f"https://trader.degiro.nl/portfolio-reports/secure/v3/positionReport/csv?intAccount=1218922&sessionId={session_id}&country=NL&lang=nl&toDate=01%2F{date.month:02}%2F{date.year}",
            output_path=os.path.join(os.getcwd(), 'data', 'exports', f'{dt.now().date().replace(day=1)}.csv')
        )
        make_export(
            page=page,
            export_link=f"https://trader.degiro.nl/portfolio-reports/secure/v3/cashAccountReport/csv?intAccount=1218922&sessionId={session_id}&country=NL&lang=nl&fromDate=27%2F05%2F1990&toDate={date.day:02}%2F{date.month:02}%2F{date.year}",
            output_path=os.path.join(os.getcwd(), 'data', 'deposits', 'Account.csv')
        )
