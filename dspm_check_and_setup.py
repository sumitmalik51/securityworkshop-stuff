"""
Combined DSPM script: checks status first, sets up DSPM if needed.
- If COMPLETED: skip (already done)
- If RUNNING: stop the previous job, then re-run setup
- If NOT_FOUND: run full DSPM Posture Agent setup
Usage: python dspm_check_and_setup.py <Sheet> <start_index> <max_count> <parallel>
"""
import time
import sys
import os
import csv
import openpyxl
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from playwright.sync_api import sync_playwright

SPN_FILE = r"C:\certs\spns.xlsx"
SPN_SHEET = sys.argv[1] if len(sys.argv) > 1 else "Batch1"
START_INDEX = int(sys.argv[2]) if len(sys.argv) > 2 else 0
MAX_COUNT = int(sys.argv[3]) if len(sys.argv) > 3 else 0
PARALLEL = int(sys.argv[4]) if len(sys.argv) > 4 else 1

RESULTS_FILE = rf"C:\certs\dspm_check_setup_{SPN_SHEET}.csv"


def load_accounts():
    accounts = []
    wb = openpyxl.load_workbook(SPN_FILE, read_only=True)
    ws = wb[SPN_SHEET]
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


def close_popups(page):
    try:
        btn = page.locator('button:has-text("Get started")')
        if btn.count() > 0 and btn.first.is_visible():
            btn.first.click(timeout=3000)
            time.sleep(2)
    except Exception:
        pass
    for _ in range(10):
        closed = False
        for sel in [
            'button[aria-label="Close"]', 'button[aria-label="Dismiss"]',
            'button:has-text("Close")', 'button:has-text("Dismiss")',
            'button:has-text("Got it")', 'button:has-text("Skip")',
            'button:has-text("OK")', 'button:has-text("No thanks")',
            'button:has-text("Maybe later")', 'button:has-text("Not now")',
            'div[role="dialog"] button[aria-label="Close"]',
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
    try:
        cp = page.locator('button[aria-label="Close Security Copilot"], button[aria-label="Close Copilot"], button[aria-label="Close panel"]')
        if cp.count() > 0 and cp.first.is_visible():
            cp.first.click(timeout=3000, force=True)
            time.sleep(1)
    except Exception:
        pass
    try:
        page.evaluate('document.querySelectorAll(".fui-DialogSurface__backdrop").forEach(e => e.remove())')
    except Exception:
        pass


def handle_setup_dialogs(page):
    clicked_get_started = False
    try:
        get_started = page.locator('button:has-text("Get started")')
        if get_started.count() > 0 and get_started.first.is_visible():
            get_started.first.click(timeout=5000)
            print("    Clicked Get started")
            clicked_get_started = True
            time.sleep(10)
    except Exception:
        pass
    try:
        start_setup = page.locator('button:has-text("Start setup")')
        if clicked_get_started:
            start_setup.first.wait_for(state="visible", timeout=15000)
        if start_setup.count() > 0 and start_setup.first.is_visible():
            start_setup.first.click(timeout=5000)
            print("    Clicked Start setup")
            time.sleep(12)
    except Exception:
        pass
    clicked_close = False
    try:
        close_btn = page.locator('button:has-text("Close")')
        if close_btn.count() > 0 and close_btn.first.is_visible():
            close_btn.first.click(timeout=5000)
            clicked_close = True
            time.sleep(3)
    except Exception:
        pass
    if clicked_close or clicked_get_started:
        try:
            page.reload(wait_until="domcontentloaded", timeout=120000)
            time.sleep(15)
        except Exception:
            pass
        return True
    return False


def login(page, username, password):
    """Login to Purview portal. Returns True on success."""
    for _login_nav in range(3):
        try:
            page.goto("https://purview.microsoft.com", timeout=120000)
            break
        except Exception:
            if _login_nav == 2:
                raise
            time.sleep(10)
    time.sleep(3)

    page.fill('input[type="email"]', username)
    page.click('input[type="submit"]')
    time.sleep(4)

    try:
        pw_link = page.locator('a:has-text("Use your password instead"), a:has-text("password"), #idA_PWD_SwitchToPassword')
        pw_link.first.click(timeout=8000)
        time.sleep(3)
    except Exception:
        pass

    try:
        page.wait_for_selector('input[type="password"]', timeout=10000)
        page.fill('input[type="password"]', password)
        page.click('input[type="submit"]')
    except Exception as e:
        print(f"    [WARN] Password issue: {e}")
    time.sleep(5)

    try:
        yes_btn = page.locator('input[type="submit"][value="Yes"], button:has-text("Yes")')
        yes_btn.first.click(timeout=15000)
    except Exception:
        pass
    time.sleep(5)

    try:
        page.wait_for_load_state("load", timeout=30000)
    except Exception:
        pass
    time.sleep(3)
    close_popups(page)


def navigate_to_dspm(page, tenant_id):
    """Navigate to DSPM and click Posture Agent. Returns True if page loaded."""
    url = f"https://purview.microsoft.com/dspm/assetexplorer?tid={tenant_id}"
    print(f"  Navigating to: {url}")
    for _nav_attempt in range(3):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=120000)
            break
        except Exception as nav_err:
            print(f"    [WARN] Navigation attempt {_nav_attempt+1} failed: {nav_err}")
            if _nav_attempt == 2:
                raise
            time.sleep(10)
    time.sleep(15)

    close_popups(page)
    handle_setup_dialogs(page)
    close_popups(page)

    # Click Posture Agent tab - poll up to 120s
    pa_clicked = False
    for _pa_wait in range(24):
        try:
            pa = page.locator('button:has-text("Posture Agent"), button:has-text("Posture agent")')
            if pa.count() > 0 and pa.first.is_visible():
                pa.first.click(timeout=5000, force=True)
                print("  Clicked Posture Agent tab")
                pa_clicked = True
                break
        except Exception:
            pass
        close_popups(page)
        handle_setup_dialogs(page)
        time.sleep(5)
    if not pa_clicked:
        print("  [WARN] Could not click Posture Agent after 120s")
    return pa_clicked


def check_dspm_status(page):
    """Check DSPM status on the Posture Agent page. Returns 'COMPLETED', 'RUNNING', or 'NOT_FOUND'."""
    # Wait for content to load (up to 120s)
    print("  Waiting for Posture Agent content to load...")
    for _content_wait in range(24):
        close_popups(page)
        handle_setup_dialogs(page)
        has_insights = page.locator('button:has-text("View insights")').count() > 0
        has_pickup = page.locator('text=/Pick up where you left off/i').count() > 0
        has_add = page.locator('button:has-text("Add data sources"), button:has-text("Add data source")').count() > 0
        has_card = page.locator('text=/Discovery agent/i, text=/Sensitive/i').count() > 0
        if has_insights or has_pickup or has_add or has_card:
            print(f"  Content loaded (insights={has_insights} pickup={has_pickup} add={has_add} card={has_card})")
            break
        time.sleep(5)
    else:
        print("  [WARN] Posture Agent content may not have loaded after 120s")

    close_popups(page)
    try:
        page.evaluate('document.querySelectorAll(".fui-DialogSurface__backdrop").forEach(e => e.remove())')
    except Exception:
        pass

    time.sleep(3)
    view_insights = page.locator('button:has-text("View insights")')
    if view_insights.count() > 0 and view_insights.first.is_visible():
        print("  STATUS: COMPLETED")
        return "COMPLETED"

    running = page.locator('text=/[Rr]unning|[Ii]n progress|[Pp]rocessing/')
    if running.count() > 0:
        print("  STATUS: RUNNING")
        return "RUNNING"

    pickup = page.locator('text=Pick up where you left off')
    if pickup.count() > 0:
        print("  STATUS: RUNNING (pick up where you left off)")
        return "RUNNING"

    print("  STATUS: NOT_FOUND")
    return "NOT_FOUND"


def stop_running_job(page):
    """Stop a running DSPM job so we can re-run it."""
    print("  >>> Stopping previous running job...")
    try:
        # Look for stop/cancel/delete buttons on the running job card
        for sel in [
            'button:has-text("Stop")',
            'button:has-text("Cancel")',
            'button:has-text("Delete")',
            'button[aria-label="Stop"]',
            'button[aria-label="Cancel"]',
            'button[aria-label="Delete"]',
        ]:
            btn = page.locator(sel)
            if btn.count() > 0 and btn.first.is_visible():
                btn.first.click(timeout=5000, force=True)
                print(f"  Clicked {sel}")
                time.sleep(3)
                break
        else:
            # Try the 3-dot / more options menu on the job card
            more = page.locator('button[aria-label="More options"], button[aria-label="More actions"], button:has-text("..."), [data-icon-name="MoreVertical"]')
            if more.count() > 0 and more.first.is_visible():
                more.first.click(timeout=5000, force=True)
                print("  Clicked More options menu")
                time.sleep(2)
                for menu_sel in [
                    'button:has-text("Stop")', 'button:has-text("Cancel")',
                    'button:has-text("Delete")', '[role="menuitem"]:has-text("Stop")',
                    '[role="menuitem"]:has-text("Cancel")', '[role="menuitem"]:has-text("Delete")',
                ]:
                    menu_btn = page.locator(menu_sel)
                    if menu_btn.count() > 0 and menu_btn.first.is_visible():
                        menu_btn.first.click(timeout=5000, force=True)
                        print(f"  Clicked menu item {menu_sel}")
                        time.sleep(3)
                        break

        # Confirm the stop/delete if a confirmation dialog appears
        for confirm_sel in [
            'button:has-text("Confirm")', 'button:has-text("Yes")',
            'button:has-text("OK")', 'button:has-text("Delete")',
            'button:has-text("Stop")',
        ]:
            try:
                confirm = page.locator(confirm_sel)
                if confirm.count() > 0 and confirm.first.is_visible():
                    confirm.first.click(timeout=5000, force=True)
                    print(f"  Confirmed with {confirm_sel}")
                    time.sleep(3)
                    break
            except Exception:
                pass

        print("  Waiting 10s after stopping job...")
        time.sleep(10)

        # Refresh the page to get clean state
        page.reload(wait_until="domcontentloaded", timeout=120000)
        time.sleep(15)
        close_popups(page)
        return True
    except Exception as e:
        print(f"  [WARN] Could not stop running job: {e}")
        # Refresh anyway to try clean state
        try:
            page.reload(wait_until="domcontentloaded", timeout=120000)
            time.sleep(15)
            close_popups(page)
        except Exception:
            pass
        return False


def run_dspm_setup(page, username, tenant_id):
    """Run the full DSPM Posture Agent setup. Returns True on success."""
    print("  >>> Running DSPM setup...")

    # Handle setup dialogs that may have appeared
    did_setup = handle_setup_dialogs(page)
    close_popups(page)
    time.sleep(2)

    # If setup happened, re-click Posture Agent
    if did_setup:
        print("  Re-clicking Posture agent after setup...")
        try:
            posture = page.locator('button:has-text("Posture Agent"), button:has-text("Posture agent")')
            if posture.count() == 0:
                posture = page.get_by_text("Posture agent", exact=False)
            posture.first.click(timeout=10000, force=True)
            time.sleep(8)
        except Exception:
            pass
        close_popups(page)
        time.sleep(2)

    # Click "Add data sources"
    print("  Clicking 'Add data sources'...")
    try:
        page.evaluate('document.querySelectorAll(".fui-DialogSurface__backdrop").forEach(e => e.remove())')
        time.sleep(1)
        add_data = page.locator('button:has-text("Add data sources"), a:has-text("Add data sources"), [role="button"]:has-text("Add data sources")')
        if add_data.count() == 0:
            add_data = page.locator('button:has-text("Add data"), a:has-text("Add data"), [role="button"]:has-text("Add data")')
        if add_data.count() == 0:
            add_data = page.get_by_text("Add data sources", exact=False)
        add_data.first.click(timeout=10000, force=True)
        print("  Clicked Add data sources")
        time.sleep(5)
    except Exception as e:
        print(f"  [WARN] Could not click Add data sources: {e}")
        try:
            page.evaluate('''() => {
                const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
                while (walker.nextNode()) {
                    if (walker.currentNode.textContent.trim().toLowerCase().includes('add data source')) {
                        walker.currentNode.parentElement.click();
                        return true;
                    }
                }
                return false;
            }''')
            time.sleep(5)
        except Exception:
            pass

    # Handle setup dialogs after Add data sources
    did_setup2 = handle_setup_dialogs(page)
    if did_setup2:
        close_popups(page)
        time.sleep(2)
        try:
            posture = page.locator('button:has-text("Posture Agent"), button:has-text("Posture agent")')
            if posture.count() == 0:
                posture = page.get_by_text("Posture agent", exact=False)
            posture.first.click(timeout=10000, force=True)
            time.sleep(8)
        except Exception:
            pass
        close_popups(page)
        time.sleep(2)
        try:
            add_data = page.locator('button:has-text("Add data sources"), a:has-text("Add data sources")')
            if add_data.count() == 0:
                add_data = page.get_by_text("Add data sources", exact=False)
            add_data.first.click(timeout=10000)
            time.sleep(5)
        except Exception:
            pass

    close_popups(page)
    time.sleep(2)

    # Search for user
    user_search = username.split("@")[0]
    print(f"  Searching for '{user_search}'...")
    try:
        search_input = page.locator('input[placeholder="Search for people, groups, or sites"]')
        search_input.fill(user_search)
        time.sleep(1)
        search_input.press("Enter")
        time.sleep(8)
    except Exception as e:
        print(f"  [WARN] Search issue: {e}")

    # Select user from results
    print("  Selecting user from results...")
    try:
        user_btn = page.locator(f'button:has-text("{user_search}")').last
        user_btn.wait_for(state="visible", timeout=15000)
        user_btn.click(timeout=5000)
        print(f"  Selected {user_search}")
        time.sleep(5)
    except Exception as e:
        print(f"  [WARN] Button click failed ({e}), trying alternatives...")
        try:
            email_el = page.locator(f'text="{username}"').last
            if email_el.count() > 0:
                email_el.click(timeout=5000)
                time.sleep(5)
            else:
                page.evaluate(f'''() => {{
                    const els = document.querySelectorAll("button, div[role='option'], div[role='listitem'], li");
                    for (const el of els) {{
                        if (el.textContent.includes("{user_search}")) {{
                            el.click();
                            return true;
                        }}
                    }}
                    return false;
                }}''')
                time.sleep(5)
        except Exception:
            pass

    handle_setup_dialogs(page)
    close_popups(page)

    # Uncheck Mailboxes, keep Sites
    print("  Unchecking Mailboxes, keeping Sites...")
    try:
        page.locator('input[type="checkbox"], [role="checkbox"]').first.wait_for(timeout=15000)
        time.sleep(2)
        checkboxes = page.locator('input[type="checkbox"], [role="checkbox"]')
        cb_count = checkboxes.count()
        print(f"  Found {cb_count} checkboxes")
        if cb_count >= 2:
            mail_cb = checkboxes.nth(0)
            if mail_cb.is_visible():
                cb_id = mail_cb.get_attribute('id')
                if cb_id:
                    label = page.locator(f'label[for="{cb_id}"]')
                    if label.count() > 0:
                        label.first.click(force=True)
                    else:
                        mail_cb.click(force=True)
                else:
                    mail_cb.click(force=True)
                print("  Unchecked Mailboxes")
        time.sleep(2)
    except Exception as e:
        print(f"  [WARN] Checkbox issue: {e}")

    # Click Save
    print("  Clicking Save...")
    try:
        save_btn = page.locator('button:has-text("Save")')
        save_btn.first.click(timeout=5000, force=True)
        print("  Clicked Save")
        time.sleep(8)
    except Exception as e:
        print(f"  [WARN] Could not click Save: {e}")

    print("  Waiting 10s after Save...")
    time.sleep(10)

    # Enter prompt and click Send
    print("  Entering prompt...")
    try:
        try:
            start_new = page.locator('button:has-text("Start new prompt")')
            if start_new.count() > 0 and start_new.first.is_visible():
                start_new.first.click(timeout=5000)
                time.sleep(5)
        except Exception:
            pass

        suggested_selectors = [
            'button:has-text("Collect files")',
            'button:has-text("Show chats")',
            'button:has-text("Retrieve financial")',
            'button:has-text("Suggested prompt")',
        ]
        clicked_suggested = False
        for sel in suggested_selectors:
            try:
                btn = page.locator(sel)
                if btn.count() > 0 and btn.first.is_visible():
                    btn.first.click(timeout=5000)
                    clicked_suggested = True
                    time.sleep(3)
                    break
            except Exception:
                continue

        if clicked_suggested:
            page.keyboard.press("Control+a")
            time.sleep(0.5)
            page.keyboard.press("Delete")
            time.sleep(0.5)
            page.keyboard.type("Get files with sensitive types password and credentials", delay=50)
            time.sleep(3)
        else:
            prompt_box = page.get_by_placeholder("Describe what you're looking for")
            prompt_box.click(timeout=10000)
            time.sleep(1)
            page.keyboard.type("Get files with sensitive types password and credentials", delay=50)
            time.sleep(3)

        send_btn = page.locator('button[aria-label="Send"], button:has-text("Send")')
        if send_btn.count() > 0:
            time.sleep(2)
            send_btn.first.click(force=True, timeout=5000)
            print("  Clicked Send")
        else:
            page.keyboard.press("Enter")
            print("  Pressed Enter to send")

        # Wait for Confirm page
        print("  Waiting for Confirm page...")
        confirm_btn = page.locator('button:has-text("Confirm")')
        confirm_btn.wait_for(state="visible", timeout=90000)
        time.sleep(2)
        confirm_btn.click(timeout=5000, force=True)
        print("  Clicked Confirm")

        print("  Waiting 30s for job to start...")
        time.sleep(30)
        refresh_btn = page.locator('button[aria-label="Refresh"], button:has-text("Refresh"), button[title="Refresh"]')
        if refresh_btn.count() > 0:
            refresh_btn.first.click(timeout=5000)
        else:
            page.reload()
        time.sleep(30)
    except Exception as e:
        print(f"  [WARN] Prompt/Confirm issue: {e}")
        return False

    print("  SETUP COMPLETE")
    return True


def process_tenant(username, password, tenant_id, index, total):
    """Check status and set up DSPM if not found."""
    print(f"\n=== [{index}/{total}] {username} ===")
    result = {
        "username": username,
        "tenant_id": tenant_id,
        "status": "ERROR",
        "action": "",
        "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context()
            page = context.new_page()

            # Login
            login(page, username, password)

            # Navigate to DSPM Posture Agent
            pa_clicked = navigate_to_dspm(page, tenant_id)

            # Check status
            status = check_dspm_status(page)
            result["status"] = status

            if status == "COMPLETED":
                result["action"] = "SKIPPED"
                print(f"  Already COMPLETED - skipping")
            elif status == "RUNNING":
                # Stop the running job and re-run setup
                print("  RUNNING detected - will stop and re-run...")
                stop_running_job(page)

                # Re-navigate to Posture Agent after stop
                navigate_to_dspm(page, tenant_id)

                setup_ok = run_dspm_setup(page, username, tenant_id)
                if setup_ok:
                    result["action"] = "RESTARTED"
                    result["status"] = "RESTARTED"
                else:
                    result["action"] = "RESTART_FAILED"
                    result["status"] = "RESTART_FAILED"
            elif status == "NOT_FOUND":
                # Run DSPM setup
                setup_ok = run_dspm_setup(page, username, tenant_id)
                if setup_ok:
                    result["action"] = "SETUP_DONE"
                    result["status"] = "SETUP_DONE"
                else:
                    result["action"] = "SETUP_FAILED"
                    result["status"] = "SETUP_FAILED"

            browser.close()

    except Exception as e:
        print(f"  ERROR: {e}")
        result["status"] = "ERROR"
        result["action"] = str(e)[:200]

    print(f"  RESULT: {result['status']} | ACTION: {result['action']}")
    return result


def save_results(results):
    fieldnames = ["username", "tenant_id", "status", "action", "checked_at"]
    with open(RESULTS_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(results)
    print(f"\nResults saved to: {RESULTS_FILE}")


def main():
    accounts = load_accounts()
    print(f"Loaded {len(accounts)} accounts from sheet '{SPN_SHEET}'")

    if START_INDEX > 0:
        accounts = accounts[START_INDEX:]
    if MAX_COUNT > 0:
        accounts = accounts[:MAX_COUNT]

    total = len(accounts)
    print(f"Processing {total} accounts (start={START_INDEX}, max={MAX_COUNT}, parallel={PARALLEL})")

    results = []
    completed = 0
    restarted = 0
    setup_done = 0
    failed = 0

    def _tally(r):
        nonlocal completed, restarted, setup_done, failed
        if r["status"] == "COMPLETED":
            completed += 1
        elif r["status"] == "RESTARTED":
            restarted += 1
        elif r["status"] == "SETUP_DONE":
            setup_done += 1
        else:
            failed += 1
        done = completed + restarted + setup_done + failed
        print(f"  Progress: {done}/{total} (Completed={completed}, Restarted={restarted}, SetupDone={setup_done}, Failed={failed})")

    if PARALLEL <= 1:
        for i, (user, pwd, tid) in enumerate(accounts):
            r = process_tenant(user, pwd, tid, i + 1 + START_INDEX, total)
            results.append(r)
            _tally(r)
            save_results(results)
    else:
        with ThreadPoolExecutor(max_workers=PARALLEL) as executor:
            futures = {}
            for i, (user, pwd, tid) in enumerate(accounts):
                f = executor.submit(process_tenant, user, pwd, tid, i + 1 + START_INDEX, total)
                futures[f] = i

            for f in as_completed(futures):
                r = f.result()
                results.append(r)
                _tally(r)
                save_results(results)

    print(f"\n{'='*60}")
    print(f"FINAL SUMMARY - {SPN_SHEET}")
    print(f"{'='*60}")
    print(f"  COMPLETED (skipped):  {completed}")
    print(f"  RESTARTED (stopped+rerun): {restarted}")
    print(f"  SETUP_DONE (new):     {setup_done}")
    print(f"  FAILED/ERROR:         {failed}")
    print(f"  TOTAL:                {len(results)}")
    print(f"\nResults saved to: {RESULTS_FILE}")


if __name__ == "__main__":
    main()
