"""
Login to Purview portal and keep browsers open until manually closed.
Runs 5 browsers at a time. Waits until all 5 are closed before starting next batch.

Usage: python purview_open_login.py Batch4
       python purview_open_login.py Batch5
"""
import time
import sys
import os
import threading
import openpyxl
from playwright.sync_api import sync_playwright

LOGIN_URL = "https://purview.microsoft.com"
SPN_FILE = r"C:\certs\spns.xlsx"
BATCH_SIZE = 2

semaphore = threading.Semaphore(BATCH_SIZE)

def load_accounts(sheet_name):
    accounts = []
    wb = openpyxl.load_workbook(SPN_FILE, read_only=True)
    ws = wb[sheet_name]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    user_col = headers.index("odluser")
    pwd_col = headers.index("odlpassword")
    tid_col = headers.index("TenantId")
    for row in ws.iter_rows(min_row=2, values_only=True):
        username = row[user_col]
        password = row[pwd_col]
        tenant_id = row[tid_col]
        if username and password and tenant_id:
            accounts.append((str(username).strip(), str(password).strip(), str(tenant_id).strip()))
    wb.close()
    return accounts


def login_and_keep_open(username, password, tenant_id, idx, total):
    """Login to Purview and keep browser open. Returns when browser is manually closed."""
    short = username.split("@")[0]
    semaphore.acquire()
    print(f"  [{idx}/{total}] {short} — launching browser...")
    try:
        pw = sync_playwright().start()
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto(LOGIN_URL)
        time.sleep(3)

        # Enter username
        page.fill('input[type="email"]', username)
        page.click('input[type="submit"]')
        time.sleep(4)

        # Click "Use your password instead"
        try:
            pw_link = page.locator('a:has-text("Use your password instead"), a:has-text("password"), #idA_PWD_SwitchToPassword')
            pw_link.first.click(timeout=8000)
            time.sleep(3)
        except Exception:
            pass

        # Enter password
        try:
            page.wait_for_selector('input[type="password"]', timeout=10000)
            page.fill('input[type="password"]', password)
            page.click('input[type="submit"]')
        except Exception as e:
            print(f"  [{idx}/{total}] {short} — password issue: {e}")

        time.sleep(4)

        # "Stay signed in?" — Yes
        try:
            yes_btn = page.locator('input[type="submit"][value="Yes"], button:has-text("Yes")')
            yes_btn.first.click(timeout=15000)
        except Exception:
            pass

        time.sleep(5)

        # Wait for page to load
        try:
            page.wait_for_load_state("load", timeout=30000)
        except Exception:
            pass

        # Click "Get started" first
        for attempt in range(10):
            try:
                gs = page.locator('button:has-text("Get started")')
                if gs.count() > 0 and gs.first.is_visible():
                    gs.first.click(timeout=3000)
                    print(f"  [{idx}/{total}] {short} — clicked Get started")
                    time.sleep(3)
                    break
            except Exception:
                pass
            time.sleep(2)

        # Dismiss remaining popups
        for _ in range(5):
            closed = False
            for sel in [
                'button:has-text("Get started")',
                'button[aria-label="Close"]',
                'button:has-text("Close")',
                'button:has-text("Dismiss")',
                'button:has-text("Got it")',
                'button:has-text("Skip")',
                'button:has-text("OK")',
                'button:has-text("Not now")',
                'button:has-text("Maybe later")',
            ]:
                try:
                    btn = page.locator(sel)
                    if btn.count() > 0 and btn.first.is_visible():
                        btn.first.click(timeout=2000)
                        closed = True
                        time.sleep(1)
                except Exception:
                    pass
            if not closed:
                break

        # Open IRM Policies tab
        irm_url = f"https://purview.microsoft.com/insiderriskmgmt/policiespage?tid={tenant_id}"
        irm_page = context.new_page()
        irm_page.goto(irm_url)
        print(f"  [{idx}/{total}] {short} — opened IRM Policies tab")
        time.sleep(3)

        # Open IRM Alerts tab
        irm_alerts_url = f"https://purview.microsoft.com/insiderriskmgmt/alertspage?tid={tenant_id}"
        irm_alerts_page = context.new_page()
        irm_alerts_page.goto(irm_alerts_url)
        print(f"  [{idx}/{total}] {short} — opened IRM Alerts tab")
        time.sleep(3)

        # Open DLP Alerts tab
        dlp_url = f"https://purview.microsoft.com/datalossprevention/alertspage?tid={tenant_id}"
        dlp_page = context.new_page()
        dlp_page.goto(dlp_url)
        print(f"  [{idx}/{total}] {short} — opened DLP tab")
        time.sleep(3)

        # Open DSPM Asset Explorer tab
        dspm_url = f"https://purview.microsoft.com/dspm/assetexplorer?tid={tenant_id}"
        dspm_page = context.new_page()
        dspm_page.goto(dspm_url)
        print(f"  [{idx}/{total}] {short} — opened DSPM tab")
        time.sleep(3)

        # Open DSI Agent Investigation tab
        dsi_url = f"https://purview.microsoft.com/dsi/agentinvestigation?tid={tenant_id}"
        dsi_page = context.new_page()
        dsi_page.goto(dsi_url)
        print(f"  [{idx}/{total}] {short} — opened DSI tab")
        time.sleep(2)

        print(f"  [{idx}/{total}] {short} — LOGGED IN (6 tabs). Close browser when done.")

        # Wait until browser is closed by probing a page
        while True:
            try:
                page.evaluate("1+1")
                time.sleep(2)
            except Exception:
                break

        print(f"  [{idx}/{total}] {short} — browser closed.")
        try:
            pw.stop()
        except Exception:
            pass
        semaphore.release()

    except Exception as e:
        print(f"  [{idx}/{total}] {short} — ERROR: {e}")
        semaphore.release()


def main():
    sheet = sys.argv[1] if len(sys.argv) > 1 else "Batch4"
    accounts = load_accounts(sheet)
    total = len(accounts)
    print(f"Loaded {total} accounts from '{sheet}'")
    print(f"Max {BATCH_SIZE} browsers open at a time. Close one → next launches automatically.\n")

    threads = []
    for i, (user, pwd, tid) in enumerate(accounts):
        t = threading.Thread(target=login_and_keep_open, args=(user, pwd, tid, i + 1, total))
        threads.append(t)
        t.start()
        time.sleep(2)  # stagger launches

    # Wait for all to finish
    for t in threads:
        t.join()

    print("\nAll accounts processed.")


if __name__ == "__main__":
    main()
