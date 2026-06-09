"""
Login to Purview, navigate to IRM Alerts, and run Triage Agent on each alert.

Usage: python purview_triage_agent.py
       Reads accounts from purview_triage_accounts.xlsx or pass inline.
"""
import time
import sys
import os
import threading
import openpyxl
from playwright.sync_api import sync_playwright

LOGIN_URL = "https://purview.microsoft.com"
SPN_FILE = r"C:\certs\spns.xlsx"
BATCH_SIZE = 5
semaphore = threading.Semaphore(BATCH_SIZE)


def login_to_purview(page, username, password, label=""):
    """Shared login flow: enter email, password, stay signed in, get started, dismiss popups."""
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
        print(f"  {label} password issue: {e}")

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

    # Click "Get started"
    for attempt in range(10):
        try:
            gs = page.locator('button:has-text("Get started")')
            if gs.count() > 0 and gs.first.is_visible():
                gs.first.click(timeout=3000)
                print(f"  {label} clicked Get started")
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


def run_triage_on_alerts(page, context, username, tenant_id):
    """Navigate to IRM Alerts page and run Triage Agent on each alert."""
    short = username.split("@")[0]
    alerts_url = f"https://purview.microsoft.com/insiderriskmgmt/alertspage?tid={tenant_id}"

    # Navigate to alerts page
    page.goto(alerts_url)
    print(f"  {short} — navigated to IRM Alerts page")
    time.sleep(8)

    # Wait for alerts table to load
    try:
        page.wait_for_load_state("networkidle", timeout=30000)
    except Exception:
        pass
    time.sleep(5)

    # Click the "Triage Agent" tab
    try:
        triage_tab = page.locator('button:has-text("Triage Agent"), [role="tab"]:has-text("Triage Agent"), span:has-text("Triage Agent")')
        if triage_tab.count() > 0:
            triage_tab.first.click(timeout=5000)
            print(f"  {short} — clicked 'Triage Agent' tab")
            time.sleep(5)
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            time.sleep(3)
        else:
            print(f"  {short} — 'Triage Agent' tab not found, continuing...")
    except Exception as e:
        print(f"  {short} — could not click Triage Agent tab: {e}")

    # Find all alert row checkboxes
    # The alerts are in a list/table with checkboxes
    # Try to find alert rows by looking for checkboxes in the alert list
    alert_checkboxes = page.locator('div[role="row"] div[role="checkbox"], div[role="row"] input[type="checkbox"], [data-automationid="DetailsRowCheck"]')
    count = alert_checkboxes.count()

    if count == 0:
        # Try alternative selectors
        alert_checkboxes = page.locator('[role="gridcell"] [role="checkbox"], .ms-DetailsRow-check, [data-selection-toggle="true"]')
        count = alert_checkboxes.count()

    if count == 0:
        # Try yet another approach - look for rows in the grid
        alert_rows = page.locator('div[role="row"][data-automationid="DetailsRow"], div[role="row"].ms-DetailsRow')
        count = alert_rows.count()
        if count == 0:
            print(f"  {short} — no alerts found on the page (0 alerts)")
            return 0
        print(f"  {short} — found {count} alert rows, will click each")
        # Use rows directly
        for i in range(count):
            _process_alert_by_row(page, alert_rows, i, count, short)
        return count

    print(f"  {short} — found {count} alerts")

    for i in range(count):
        _process_alert_by_checkbox(page, alert_checkboxes, i, count, short)

    return count


def _process_alert_by_checkbox(page, checkboxes, idx, total, short):
    """Select an alert by checkbox, run agent, wait for completion."""
    alert_num = idx + 1
    print(f"  {short} — processing alert {alert_num}/{total}...")

    try:
        # Click the checkbox to select this alert
        cb = checkboxes.nth(idx)
        cb.scroll_into_view_if_needed()
        time.sleep(1)
        cb.click(timeout=5000)
        time.sleep(2)
        print(f"  {short} — selected alert {alert_num}/{total}")

        # Click "Run agent" button
        _click_run_agent(page, short, alert_num, total)

        # Uncheck the alert after agent finishes
        try:
            # Click the "X selected" dismiss or uncheck
            dismiss = page.locator('button:has-text("selected"), span:has-text("selected")')
            if dismiss.count() > 0:
                # Click the X near "1 selected" to deselect
                x_btn = page.locator('[aria-label="Clear selection"], button:has-text("×")')
                if x_btn.count() > 0:
                    x_btn.first.click(timeout=3000)
                else:
                    cb.click(timeout=3000)  # toggle off
            else:
                cb.click(timeout=3000)
            time.sleep(1)
        except Exception:
            pass

    except Exception as e:
        print(f"  {short} — error on alert {alert_num}: {e}")


def _process_alert_by_row(page, rows, idx, total, short):
    """Select an alert by clicking its row checkbox, run agent, wait for completion."""
    alert_num = idx + 1
    print(f"  {short} — processing alert {alert_num}/{total}...")

    try:
        row = rows.nth(idx)
        row.scroll_into_view_if_needed()
        time.sleep(1)

        # Find checkbox within the row
        cb = row.locator('[role="checkbox"], [data-automationid="DetailsRowCheck"], .ms-DetailsRow-check')
        if cb.count() > 0:
            cb.first.click(timeout=5000)
        else:
            # Click the row itself to select
            row.click(timeout=5000)
        time.sleep(2)
        print(f"  {short} — selected alert {alert_num}/{total}")

        # Click "Run agent" button
        _click_run_agent(page, short, alert_num, total)

        # Deselect
        try:
            if cb.count() > 0:
                cb.first.click(timeout=3000)
            time.sleep(1)
        except Exception:
            pass

    except Exception as e:
        print(f"  {short} — error on alert {alert_num}: {e}")


def _click_run_agent(page, short, alert_num, total):
    """Click 'Run agent' and wait for it to finish."""
    try:
        # Look for "Run agent" button with multiple selector strategies
        run_btn = page.locator('button:has-text("Run agent")')
        if run_btn.count() == 0:
            run_btn = page.locator('[aria-label="Run agent"]')
        if run_btn.count() == 0:
            run_btn = page.locator('text="Run agent"')
        if run_btn.count() == 0:
            # Try finding by icon + text in command bar
            run_btn = page.locator('[data-automationid="splitbuttonprimary"]:has-text("Run agent"), [role="menuitem"]:has-text("Run agent")')
        if run_btn.count() == 0:
            # Broader: any clickable element containing "Run agent"
            run_btn = page.locator('*:has-text("Run agent")').last

        if run_btn.count() > 0:
            run_btn.first.click(timeout=5000)
            print(f"  {short} — clicked 'Run agent' for alert {alert_num}/{total}")
        else:
            print(f"  {short} — 'Run agent' button not found for alert {alert_num}/{total}")
            return

        # Handle confirmation popup — click "Run" button in the dialog
        time.sleep(3)
        for sel in [
            'button:has-text("Run")',
            '[role="dialog"] button:has-text("Run")',
            'button:has-text("Confirm")',
            'button:has-text("Yes")',
            '[role="dialog"] button[class*="primary"]',
        ]:
            try:
                confirm_btn = page.locator(sel)
                if confirm_btn.count() > 0 and confirm_btn.first.is_visible():
                    confirm_btn.first.click(timeout=5000)
                    print(f"  {short} — clicked confirmation 'Run' for alert {alert_num}/{total}")
                    break
            except Exception:
                pass
        time.sleep(2)

        # Wait for agent to finish processing (up to 3 minutes)
        # The agent runs and we need to detect when it's done
        # Look for status changes, loading spinners, or result indicators
        time.sleep(5)

        max_wait = 180  # 3 minutes max
        waited = 0
        poll_interval = 5

        while waited < max_wait:
            # Check if a loading/processing indicator is visible
            processing = False
            for indicator in [
                '[role="progressbar"]',
                '.ms-Spinner',
                'div:has-text("Running")',
                'div:has-text("Processing")',
                'div:has-text("Analyzing")',
                '.agent-running',
                '[data-automationid="spinner"]',
            ]:
                try:
                    el = page.locator(indicator)
                    if el.count() > 0 and el.first.is_visible():
                        processing = True
                        break
                except Exception:
                    pass

            if not processing:
                # Check if result/summary appeared
                # Look for categorization or triage result
                for result_sel in [
                    'div:has-text("Needs attention")',
                    'div:has-text("Less urgent")',
                    'div:has-text("Not categorized")',
                    'div:has-text("Agent completed")',
                    'div:has-text("Triage complete")',
                    'button:has-text("Confirm")',
                    'button:has-text("Dismiss")',
                ]:
                    try:
                        r = page.locator(result_sel)
                        if r.count() > 0 and r.first.is_visible():
                            # Looks like agent finished
                            break
                    except Exception:
                        pass

                # If we've waited at least 10 seconds and no spinner, likely done
                if waited >= 10:
                    break

            time.sleep(poll_interval)
            waited += poll_interval

        # Dismiss any result dialogs/panels
        time.sleep(2)
        for sel in [
            'button:has-text("Close")',
            'button:has-text("Dismiss")',
            'button:has-text("Got it")',
            'button[aria-label="Close"]',
        ]:
            try:
                btn = page.locator(sel)
                if btn.count() > 0 and btn.first.is_visible():
                    btn.first.click(timeout=2000)
                    time.sleep(1)
            except Exception:
                pass

        print(f"  {short} — agent finished for alert {alert_num}/{total} (waited {waited}s)")

    except Exception as e:
        print(f"  {short} — error running agent for alert {alert_num}: {e}")


def process_account(username, password, tenant_id, idx=1, total=1, keep_open=True):
    """Full flow: login, navigate to alerts, run triage agent on each."""
    short = username.split("@")[0]
    if not keep_open:
        semaphore.acquire()
    print(f"\n[{idx}/{total}] {short} — starting...")

    try:
        pw = sync_playwright().start()
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        label = f"[{idx}/{total}] {short}"

        # Login
        login_to_purview(page, username, password, label)
        print(f"  {short} — logged in successfully")

        # Run triage on all alerts
        alert_count = run_triage_on_alerts(page, context, username, tenant_id)
        print(f"  {short} — processed {alert_count} alerts total")

        if keep_open:
            # Keep browser open for review
            print(f"  {short} — DONE. Close browser when ready.")
            while True:
                try:
                    page.evaluate("1+1")
                    time.sleep(2)
                except Exception:
                    break
            print(f"  {short} — browser closed.")
        else:
            print(f"  {short} — DONE. Closing browser.")
            try:
                browser.close()
            except Exception:
                pass

        try:
            pw.stop()
        except Exception:
            pass

    except Exception as e:
        print(f"  {short} — ERROR: {e}")
    finally:
        if not keep_open:
            semaphore.release()


def main():
    # Detect mode: if first arg contains "@", it's a username (single account mode)
    if len(sys.argv) >= 4 and "@" in sys.argv[1]:
        username = sys.argv[1]
        password = sys.argv[2]
        tenant_id = sys.argv[3]
        process_account(username, password, tenant_id)
    elif len(sys.argv) >= 2:
        # Sheet mode from Excel — supports multiple sheets
        sheets = sys.argv[1:]
        accounts = []
        wb = openpyxl.load_workbook(SPN_FILE, read_only=True)
        for sheet in sheets:
            ws = wb[sheet]
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            user_col = headers.index("odluser")
            pwd_col = headers.index("odlpassword")
            tid_col = headers.index("TenantId")
            for row in ws.iter_rows(min_row=2, values_only=True):
                u, p, t = row[user_col], row[pwd_col], row[tid_col]
                if u and p and t:
                    accounts.append((str(u).strip(), str(p).strip(), str(t).strip()))
        wb.close()
        total = len(accounts)
        print(f"Loaded {total} accounts from {sheets}")
        print(f"Max {BATCH_SIZE} browsers at a time.\n")
        threads = []
        for i, (u, p, t) in enumerate(accounts):
            t_thread = threading.Thread(target=process_account, args=(u, p, t, i + 1, total, False))
            threads.append(t_thread)
            t_thread.start()
            time.sleep(3)  # stagger launches
        for t_thread in threads:
            t_thread.join()
        print("\nAll accounts processed.")
    else:
        print("Usage: python purview_triage_agent.py <username> <password> <tenant_id>")
        print("       python purview_triage_agent.py <SheetName> [SheetName2] ...")


if __name__ == "__main__":
    main()
