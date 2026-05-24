import time
import sys
import os
import threading
import openpyxl
from playwright.sync_api import sync_playwright
 
LOGIN_URL = "https://purview.microsoft.com"
SPN_FILE = r"C:\certs\spns.xlsx"
SPN_SHEET = sys.argv[1] if len(sys.argv) > 1 else "All"

def load_accounts():
    """Load accounts from spns.xlsx sheet (odluser + odlpassword + TenantId)."""
    accounts = []
    wb = openpyxl.load_workbook(SPN_FILE, read_only=True)
    ws = wb[SPN_SHEET]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    user_col = headers.index("odluser")
    pwd_col = headers.index("odlpassword")
    tid_col = headers.index("TenantId")
    row_idx = 1  # 0-based data row index (for writing back)
    for row in ws.iter_rows(min_row=2, values_only=True):
        username = row[user_col]
        password = row[pwd_col]
        tenant_id = row[tid_col]
        if username and password and tenant_id:
            accounts.append((str(username).strip(), str(password).strip(), str(tenant_id).strip(), row_idx))
        row_idx += 1
    wb.close()
    return accounts

ACCOUNTS = load_accounts()
print(f"Loaded {len(ACCOUNTS)} accounts from {SPN_FILE} (sheet: {SPN_SHEET}).")

results_lock = threading.Lock()
results = []  # list of {username, status, error}


def login_account(username, password, tenant_id, row_idx=0):
    error_msg = None
    nav_solutions = "NO"
    nav_irm = "NO"
    nav_policies = "NO"
    try:
      with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
 
        page.goto(LOGIN_URL)
        time.sleep(3)
 
        # Enter username
        page.fill('input[type="email"]', username)
        page.click('input[type="submit"]')
        time.sleep(4)
 
        # Click "Password" option if "tap" / other method is shown
        try:
            password_link = page.locator('a:has-text("Use your password instead"), a:has-text("password"), #idA_PWD_SwitchToPassword')
            password_link.first.click(timeout=8000)
            time.sleep(3)
        except Exception:
            pass  # Already on password page
 
        # Enter password
        try:
            page.wait_for_selector('input[type="password"]', timeout=10000)
            page.fill('input[type="password"]', password)
            page.click('input[type="submit"]')
        except Exception as e:
            print(f"     [WARN] Password step issue: {e}")
        time.sleep(4)
 
        # Click "Yes" on Stay signed in prompt
        try:
            yes_button = page.locator('input[type="submit"][value="Yes"], button:has-text("Yes")')
            yes_button.first.click(timeout=15000)
            time.sleep(2)
        except Exception:
            pass
 
        # Wait for Purview portal to load
        try:
            page.wait_for_load_state("load", timeout=30000)
        except Exception:
            pass
        time.sleep(8)
 
        # Step 1: Click "Get started" on Welcome popup
        try:
            get_started = page.locator('button:has-text("Get started")')
            if get_started.count() > 0 and get_started.first.is_visible():
                get_started.first.click(timeout=5000)
                print("     [POPUP] Clicked 'Get started'")
                time.sleep(3)
        except Exception:
            pass

        # Step 2: Close tour popups ("Easy access to solutions" etc.) by clicking X
        for attempt in range(10):
            closed_something = False
            for selector in [
                'button[aria-label="Close"]',
                'button[aria-label="Dismiss"]',
                'button:has-text("Close")',
                'button:has-text("Dismiss")',
                'button:has-text("Got it")',
                'button:has-text("Skip")',
                'button:has-text("OK")',
                'button:has-text("No thanks")',
                'button:has-text("Maybe later")',
                'button:has-text("Not now")',
                '.ms-Dialog-main button.ms-Dialog-button--close',
                '[data-automationid="dismissButton"]',
                'div[role="dialog"] button[aria-label="Close"]',
            ]:
                try:
                    btn = page.locator(selector)
                    if btn.count() > 0 and btn.first.is_visible():
                        btn.first.click(timeout=2000)
                        closed_something = True
                        print(f"     [POPUP] Closed: {selector}")
                        time.sleep(2)
                except Exception:
                    pass
            if not closed_something:
                break
            time.sleep(1)

        # Step 3: Close Security Copilot side panel (top-right X)
        try:
            copilot_close = page.locator('button[aria-label="Close Security Copilot"], button[aria-label="Close Copilot"], button[aria-label="Close panel"]')
            if copilot_close.count() > 0 and copilot_close.first.is_visible():
                copilot_close.first.click(timeout=3000, force=True)
                print("     [POPUP] Closed Security Copilot panel")
                time.sleep(2)
        except Exception:
            pass

        # Remove any remaining dialog backdrops that block clicks
        try:
            page.evaluate('document.querySelectorAll(".fui-DialogSurface__backdrop").forEach(e => e.remove())')
            page.evaluate('document.querySelectorAll("[aria-hidden=true]").forEach(e => { if(e.classList.contains("fui-DialogSurface__backdrop") || e.classList.contains("r1e18s3l")) e.remove() })')
        except Exception:
            pass

        time.sleep(2)

        # Navigate directly to IRM Policies page using tenant-specific URL
        irm_url = f"https://purview.microsoft.com/insiderriskmgmt/policiespage?tid={tenant_id}"
        print(f"     Navigating to IRM Policies: {irm_url}")
        try:
            page.goto(irm_url, wait_until="domcontentloaded", timeout=60000)
            nav_solutions = "YES"
            time.sleep(10)
        except Exception as e:
            print(f"     [WARN] Could not navigate to IRM Policies: {e}")

        # Close any popups that appear
        for attempt in range(3):
            closed_something = False
            for selector in [
                'button[aria-label="Close"]',
                'button:has-text("Close")',
                'button:has-text("Dismiss")',
                'button:has-text("Got it")',
                'button:has-text("Get started")',
                'button:has-text("OK")',
                'button:has-text("Skip")',
                'button:has-text("Not now")',
                'button:has-text("Maybe later")',
                'div[role="dialog"] button[aria-label="Close"]',
            ]:
                try:
                    btn = page.locator(selector)
                    if btn.count() > 0 and btn.first.is_visible():
                        btn.first.click(timeout=2000, force=True)
                        closed_something = True
                        print(f"     [POPUP] Closed: {selector}")
                        time.sleep(2)
                except Exception:
                    pass
            if not closed_something:
                break
            time.sleep(1)

        # Verify URL
        current_url = page.url
        print(f"     Current URL: {current_url}")
        if "insiderriskmgmt" in current_url.lower() or "policiespage" in current_url.lower():
            nav_irm = "YES"
            nav_policies = "YES"
            print("     Confirmed: on IRM Policies page")
        else:
            print(f"     [WARN] May not be on IRM Policies page")

        # Check for permission/error messages
        try:
            error_texts = page.locator('text=/don.*have.*permission|not authorized|access denied|not available|something went wrong/i')
            if error_texts.count() > 0:
                msg = error_texts.first.text_content(timeout=2000)
                nav_policies = f"ERROR: {msg[:80]}"
                print(f"     [WARN] Page error: {msg[:80]}")
        except Exception:
            pass

        # Click on Overview in left nav
        print("     Clicking Overview...")
        try:
            overview = page.locator('a:has-text("Overview"), button:has-text("Overview"), [aria-label="Overview"]')
            if overview.count() > 0:
                overview.first.click(timeout=10000, force=True)
                print("     Clicked Overview")
                time.sleep(5)
                # Wait for recommended actions to load
                try:
                    page.wait_for_selector('text=recommended actions', timeout=15000)
                    time.sleep(3)
                except Exception:
                    time.sleep(5)
        except Exception as e:
            print(f"     [WARN] Could not click Overview: {e}")

        # Click all recommended actions
        recommended_actions = [
            "Turn on analytics",
            "Get to know insider risk",
            "Configure insider risk settings",
            "Create your first policy",
            "Make sure your team can get their jobs done",
        ]
        for action_text in recommended_actions:
            print(f"     Clicking '{action_text}'...")
            try:
                # Use JavaScript to find and click elements containing the text
                clicked = page.evaluate('''(text) => {
                    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
                    while (walker.nextNode()) {
                        if (walker.currentNode.textContent.includes(text)) {
                            let el = walker.currentNode.parentElement;
                            if (el && el.offsetHeight > 0) {
                                el.click();
                                return true;
                            }
                        }
                    }
                    return false;
                }''', action_text)
                if clicked:
                    print(f"     Clicked '{action_text}'")
                    time.sleep(3)
                    # Handle any confirmation dialog
                    for sel in ['button:has-text("Turn on")', 'button:has-text("Enable")', 'button:has-text("Yes")', 'button:has-text("Confirm")', 'button:has-text("Save")', 'button:has-text("Done")']:
                        try:
                            btn = page.locator(sel)
                            if btn.count() > 0 and btn.first.is_visible():
                                btn.first.click(timeout=5000, force=True)
                                print(f"     Confirmed: {sel}")
                                time.sleep(2)
                                break
                        except Exception:
                            pass
                    # Go back to Overview if navigated away
                    if "overview" not in page.url.lower():
                        page.go_back()
                        time.sleep(3)
                else:
                    print(f"     [INFO] '{action_text}' not found (may already be done)")
            except Exception as e:
                print(f"     [WARN] Could not click '{action_text}': {e}")

        # Close any final popups
        for selector in ['button[aria-label="Close"]', 'button:has-text("Close")', 'button:has-text("Dismiss")', 'button:has-text("Got it")']:
            try:
                btn = page.locator(selector)
                if btn.count() > 0 and btn.first.is_visible():
                    btn.first.click(timeout=2000)
                    time.sleep(1)
            except Exception:
                pass
 
        print(f"[OK] Logged in: {username}")
        print("     Waiting 30 seconds before closing browser...")
        time.sleep(30)

        # Close browser automatically
        try:
            browser.close()
        except Exception:
            pass
 
        print("     Browser closed. Moving on...")
    except Exception as e:
        error_msg = str(e)
        print(f"[FAIL] {username}: {error_msg}")
        try:
            browser.close()
        except Exception:
            pass

    with results_lock:
        results.append({
            "username": username,
            "row_idx": row_idx,
            "status": "OK" if error_msg is None else "FAIL",
            "solutions": nav_solutions,
            "irm": nav_irm,
            "policies": nav_policies,
            "error": error_msg or ""
        })
 
 
def main():
    accounts = ACCOUNTS
    batch_size = 3  # Run 3 at a time
    print(f"Found {len(accounts)} accounts. Running in batches of {batch_size}.\n")

    for batch_start in range(0, len(accounts), batch_size):
        batch = accounts[batch_start:batch_start + batch_size]
        batch_num = batch_start // batch_size + 1
        total_batches = (len(accounts) + batch_size - 1) // batch_size
        print(f"\n===== Batch {batch_num}/{total_batches} ({len(batch)} accounts) =====")

        threads = []
        for i, (username, password, tenant_id, row_idx) in enumerate(batch, batch_start + 1):
            t = threading.Thread(target=login_account, args=(username, password, tenant_id, row_idx), name=f"Account-{i}")
            threads.append(t)

        for t in threads:
            t.start()
            time.sleep(2)

        for t in threads:
            t.join()

        print(f"===== Batch {batch_num} complete =====")

    # Write results back into spns.xlsx (with file lock for parallel safety)
    lock_file = SPN_FILE + ".lock"
    print("Acquiring file lock for writing results...")
    while True:
        try:
            fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            break
        except FileExistsError:
            print("  File locked by another batch, waiting 3s...")
            time.sleep(3)

    saved = False
    try:
        wb = openpyxl.load_workbook(SPN_FILE)
        ws = wb[SPN_SHEET]
        headers = [cell.value for cell in ws[1]]

        # Add new columns if they don't exist
        new_cols = ["purview_status", "purview_solutions", "purview_irm", "purview_policies", "purview_error"]
        for col_name in new_cols:
            if col_name not in headers:
                headers.append(col_name)
                ws.cell(row=1, column=len(headers), value=col_name)

        # Write each result to its row
        for r in results:
            data_row = r["row_idx"] + 1  # +1 for header row -> Excel row
            ws.cell(row=data_row + 1, column=headers.index("purview_status") + 1, value=r["status"])
            ws.cell(row=data_row + 1, column=headers.index("purview_solutions") + 1, value=r["solutions"])
            ws.cell(row=data_row + 1, column=headers.index("purview_irm") + 1, value=r["irm"])
            ws.cell(row=data_row + 1, column=headers.index("purview_policies") + 1, value=r["policies"])
            ws.cell(row=data_row + 1, column=headers.index("purview_error") + 1, value=r["error"])

        # Retry save up to 5 times if file is locked by Excel
        for attempt in range(5):
            try:
                wb.save(SPN_FILE)
                saved = True
                break
            except PermissionError:
                print(f"  [WARN] spns.xlsx is locked (attempt {attempt+1}/5). Close Excel and waiting 10s...")
                time.sleep(10)
        wb.close()
    finally:
        if os.path.exists(lock_file):
            os.remove(lock_file)

    if saved:
        print(f"\nResults written back to {SPN_FILE} (sheet: {SPN_SHEET})")
    else:
        # Fallback: save results to CSV
        import csv
        csv_file = SPN_FILE.replace(".xlsx", f"_purview_{SPN_SHEET}.csv")
        with open(csv_file, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["username", "row_idx", "status", "solutions", "irm", "policies", "error"])
            writer.writeheader()
            writer.writerows(results)
        print(f"\n[WARN] Could not save to spns.xlsx (file locked). Results saved to: {csv_file}")

    failures = [r for r in results if r["status"] == "FAIL"]
    nav_issues = [r for r in results if r["status"] == "OK" and (r["solutions"] == "NO" or r["irm"] == "NO" or r["policies"] == "NO")]
    print(f"All {len(results)} accounts processed. {len(failures)} login failures. {len(nav_issues)} with navigation issues.")
    if failures:
        print("\nFailed logins:")
        for f in failures:
            print(f"  - {f['username']}: {f['error']}")
    if nav_issues:
        print(f"\nNavigation issues ({len(nav_issues)} accounts):")
        for r in nav_issues:
            missing = []
            if r["solutions"] == "NO": missing.append("Solutions")
            if r["irm"] == "NO": missing.append("IRM")
            if r["policies"] == "NO": missing.append("Policies")
            print(f"  - {r['username']}: missing {', '.join(missing)}")
 
 
if __name__ == "__main__":
    main()
 