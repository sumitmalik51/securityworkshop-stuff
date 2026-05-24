"""Quick debug script to inspect the scoring dialog DOM elements."""
import time
import sys
import openpyxl
from playwright.sync_api import sync_playwright

SPN_FILE = r"C:\certs\spns.xlsx"

wb = openpyxl.load_workbook(SPN_FILE, read_only=True)
ws = wb["Batch1"]
headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
user_col = headers.index("odluser")
pwd_col = headers.index("odlpassword")
tid_col = headers.index("TenantId")
rows = list(ws.iter_rows(min_row=2, values_only=True))
wb.close()

row = rows[13]  # odl_user_2226972
username = str(row[user_col]).strip()
password = str(row[pwd_col]).strip()
tenant_id = str(row[tid_col]).strip()

print(f"User: {username}")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    # Login
    page.goto("https://purview.microsoft.com")
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
    page.fill('input[type="password"]', password)
    page.click('input[type="submit"]')
    time.sleep(5)
    try:
        yes_btn = page.locator('input[type="submit"][value="Yes"], button:has-text("Yes")')
        yes_btn.first.click(timeout=8000)
        time.sleep(3)
    except Exception:
        pass

    # Navigate to IRM Policies
    url = f"https://purview.microsoft.com/insiderriskmgmt/policiespage?tid={tenant_id}"
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    time.sleep(5)

    # Dismiss "Switch to classic portal" dialog first
    for attempt in range(5):
        try:
            dialogs = page.locator('[role="dialog"]')
            for i in range(dialogs.count()):
                txt = dialogs.nth(i).text_content()
                if txt and 'Switch back' in txt:
                    close_btn = dialogs.nth(i).locator('button[aria-label="Close"], button:has-text("Close")')
                    if close_btn.count() > 0:
                        close_btn.first.click(timeout=3000, force=True)
                        print(f"Dismissed Switch dialog (attempt {attempt+1})")
                        time.sleep(2)
        except Exception:
            pass
        # Also try Escape
        page.keyboard.press("Escape")
        time.sleep(1)

    # Close other popups
    for _ in range(5):
        closed = False
        for sel in [
            'button[aria-label="Close"]', 'button[aria-label="Dismiss"]',
            'button:has-text("Got it")', 'button:has-text("Skip")',
            'button:has-text("OK")', 'button:has-text("No thanks")',
            'button:has-text("Maybe later")', 'button:has-text("Not now")',
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

    # Wait for policies to load (shimmers to disappear)
    print("Waiting for policies to load (shimmers)...")
    for wait_attempt in range(30):  # up to 60 seconds
        shimmer_count = page.locator('.ms-Shimmer-container').count()
        row_count = page.locator('.ms-DetailsRow').count()
        print(f"  wait {wait_attempt*2}s: shimmers={shimmer_count}, rows={row_count}")
        if row_count > 0:
            print(f"  Policies loaded! Found {row_count} rows")
            break
        time.sleep(2)
    else:
        print("  Policies did not load within 60 seconds!")

    time.sleep(3)

    # Dump page structure to understand what loaded
    print("\n=== PAGE STRUCTURE ===")
    page_struct = page.evaluate('''() => {
        const result = [];
        // Look for DetailsList rows (ms-DetailsRow inside ListCell)
        document.querySelectorAll('.ms-DetailsRow, [data-automationid="DetailsRow"]').forEach((el, idx) => {
            const check = el.querySelector('.ms-DetailsRow-check, [data-automationid="DetailsRowCheck"]');
            result.push({
                type: 'row',
                index: idx,
                selIndex: el.getAttribute('data-selection-index'),
                isSelected: el.classList.contains('is-selected'),
                ariaSelected: el.getAttribute('aria-selected'),
                hasCheck: !!check,
                checkRole: check ? check.getAttribute('role') : null,
                text: el.textContent.substring(0, 120),
                visible: el.offsetParent !== null
            });
        });
        // If no DetailsRow found, look at ListCell content
        if (result.length === 0) {
            document.querySelectorAll('[data-automationid="ListCell"]').forEach((el, idx) => {
                result.push({
                    type: 'listcell',
                    index: idx,
                    innerHTML: el.innerHTML.substring(0, 300),
                    text: el.textContent.substring(0, 120),
                    visible: el.offsetParent !== null
                });
            });
        }
        return result;
    }''')
    for s in page_struct:
        if s.get('visible', True):
            if s['type'] == 'row':
                print(f"  ROW idx={s['index']} selIdx={s['selIndex']} selected={s['isSelected']} ariaSelected={s['ariaSelected']} hasCheck={s['hasCheck']} checkRole={s['checkRole']}")
                print(f"    text: {s['text'][:100]}")
            else:
                print(f"  LISTCELL idx={s['index']} text={s['text'][:100]}")
                print(f"    html: {s.get('innerHTML','')[:200]}")

    # First, select all policies by clicking on each row's checkbox
    print("\n=== SELECTING POLICIES ===")
    # Dump all checkbox/check elements
    checks_info = page.evaluate('''() => {
        const result = [];
        document.querySelectorAll('.ms-Check, [role="checkbox"], [data-selection-toggle], input[type="checkbox"]').forEach(el => {
            result.push({
                tag: el.tagName,
                role: el.getAttribute('role'),
                className: el.className.substring(0, 100),
                ariaLabel: el.getAttribute('aria-label'),
                dataToggle: el.getAttribute('data-selection-toggle'),
                visible: el.offsetParent !== null,
                text: el.textContent.substring(0, 50)
            });
        });
        return result;
    }''')
    for c in checks_info:
        if c['visible']:
            print(f"  CHECK: tag={c['tag']} role={c['role']} class={c['className'][:50]} aria={c['ariaLabel']} toggle={c['dataToggle']}")

    # Try clicking the select-all check
    try:
        # Click the role=checkbox element directly (header select all)
        select_all = page.locator('[role="checkbox"][data-selection-toggle="true"]')
        if select_all.count() > 0:
            print(f"  Found {select_all.count()} select-all checkbox(es)")
            select_all.first.click(timeout=5000, force=True)
            print("  Clicked select-all checkbox")
            time.sleep(3)
        else:
            # Fallback: click .ms-Check
            select_all = page.locator('.ms-Check')
            if select_all.count() > 0:
                print(f"  Found {select_all.count()} .ms-Check elements, clicking first")
                select_all.first.click(timeout=5000, force=True)
                time.sleep(3)
    except Exception as e:
        print(f"  Error selecting: {e}")

    # Check if rows are selected
    selected_rows = page.evaluate('''() => {
        const result = [];
        document.querySelectorAll('.ms-DetailsRow, [data-selection-index]').forEach(el => {
            result.push({
                index: el.getAttribute('data-selection-index'),
                selected: el.classList.contains('is-selected'),
                ariaSelected: el.getAttribute('aria-selected'),
                text: el.textContent.substring(0, 80)
            });
        });
        return result;
    }''')
    print(f"\n=== SELECTED ROWS ({len(selected_rows)} total) ===")
    for r in selected_rows:
        print(f"  idx={r['index']} selected={r['selected']} aria-selected={r['ariaSelected']} text={r['text'][:60]}")

    # Click Start scoring activity for users
    try:
        btn = page.locator('button:has-text("Start scoring activity for users")')
        btn.first.click(timeout=10000, force=True)
        print("\nClicked Start scoring activity for users")
        time.sleep(10)  # Wait longer for panel to open
    except Exception as e:
        print(f"Error: {e}")

    # Check for panel/dialog/flyout
    panel_info = page.evaluate('''() => {
        const result = [];
        document.querySelectorAll('.ms-Panel, .ms-Dialog, [role="dialog"], [role="complementary"], .ms-Panel-main, [class*="panel" i], [class*="flyout" i], [class*="drawer" i]').forEach(el => {
            result.push({
                tag: el.tagName,
                role: el.getAttribute('role'),
                className: el.className.substring(0, 120),
                visible: el.offsetParent !== null || el.style.display !== 'none',
                childCount: el.children.length,
                text: el.textContent.substring(0, 200)
            });
        });
        return result;
    }''')
    print(f"\n=== PANELS/DIALOGS ({len(panel_info)} found) ===")
    for pi in panel_info:
        print(f"  {pi['tag']} role={pi['role']} visible={pi['visible']} children={pi['childCount']} class={pi['className'][:80]}")
        if pi['visible']:
            print(f"    text: {pi['text'][:200]}")

    # Dump all visible inputs
    inputs = page.evaluate('''() => {
        const result = [];
        document.querySelectorAll('input, textarea, select').forEach(el => {
            result.push({
                tag: el.tagName,
                type: el.type,
                placeholder: el.placeholder,
                ariaLabel: el.getAttribute('aria-label'),
                name: el.name,
                id: el.id,
                className: el.className.substring(0, 100),
                visible: el.offsetParent !== null,
                value: el.value
            });
        });
        return result;
    }''')
    
    print("\n=== ALL INPUTS IN PAGE ===")
    for inp in inputs:
        if inp['visible']:
            print(f"  [{inp['tag']}] type={inp['type']} placeholder='{inp['placeholder']}' aria-label='{inp['ariaLabel']}' id='{inp['id']}' class='{inp['className'][:60]}'")

    # Also dump buttons
    buttons = page.evaluate('''() => {
        const result = [];
        document.querySelectorAll('button').forEach(el => {
            if (el.offsetParent !== null && el.textContent.trim().length > 0 && el.textContent.trim().length < 100) {
                result.push({
                    text: el.textContent.trim().substring(0, 60),
                    ariaLabel: el.getAttribute('aria-label'),
                    className: el.className.substring(0, 80)
                });
            }
        });
        return result;
    }''')
    
    print("\n=== VISIBLE BUTTONS ===")
    for b in buttons:
        print(f"  '{b['text']}' aria-label='{b['ariaLabel']}'")

    # Keep browser open for 30s so we can look
    time.sleep(30)
    browser.close()
