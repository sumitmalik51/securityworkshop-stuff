import time, openpyxl
from playwright.sync_api import sync_playwright

wb = openpyxl.load_workbook('C:/certs/spns.xlsx', read_only=True)
ws = wb['Batch1']
headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
rows = list(ws.iter_rows(min_row=2, values_only=True))
row = rows[52]
user = str(row[headers.index('odluser')]).strip()
pwd = str(row[headers.index('odlpassword')]).strip()
tid = str(row[headers.index('TenantId')]).strip()
wb.close()
print(f'User: {user}, TID: {tid}')

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_context().new_page()
    page.goto('https://purview.microsoft.com')
    time.sleep(3)
    page.fill('input[type="email"]', user)
    page.click('input[type="submit"]')
    time.sleep(4)
    try:
        page.locator('a:has-text("Use your password instead"), #idA_PWD_SwitchToPassword').first.click(timeout=8000)
        time.sleep(3)
    except:
        pass
    page.wait_for_selector('input[type="password"]', timeout=10000)
    page.fill('input[type="password"]', pwd)
    page.click('input[type="submit"]')
    time.sleep(4)
    try:
        page.locator('input[type="submit"][value="Yes"]').first.click(timeout=15000)
    except:
        pass
    time.sleep(8)
    for sel in ['button:has-text("Get started")', 'button[aria-label="Close"]', 'button:has-text("Close")']:
        try:
            b = page.locator(sel)
            if b.count() > 0 and b.first.is_visible():
                b.first.click(timeout=2000)
                time.sleep(1)
        except:
            pass
    page.goto(f'https://purview.microsoft.com/dspm/assetexplorer?tid={tid}', wait_until='domcontentloaded', timeout=60000)
    time.sleep(10)
    # Click Posture agent if available
    try:
        pa = page.locator('button:has-text("Posture Agent")')
        if pa.count() > 0:
            pa.first.click(timeout=5000)
            print("Clicked Posture Agent button")
        else:
            page.locator('text="Posture agent (preview)"').first.click(timeout=5000)
            print("Clicked Posture agent text")
        time.sleep(8)
    except:
        print("Posture agent tab not found or already active")
        time.sleep(3)

    # Also close any popups
    for sel in ['button:has-text("Get started")', 'button[aria-label="Close"]', 'button:has-text("Close")', 'button:has-text("Got it")', 'button:has-text("Dismiss")']:
        try:
            b = page.locator(sel)
            if b.count() > 0 and b.first.is_visible():
                b.first.click(timeout=2000)
                time.sleep(1)
        except:
            pass
    time.sleep(3)

    # Take screenshot to see actual state
    page.screenshot(path="C:/certs/debug_screenshot.png")
    print("Screenshot saved to debug_screenshot.png")

    # Check if Data Security Posture Agent text is visible
    dspa = page.get_by_text("Data Security Posture Agent")
    print(f"\n'Data Security Posture Agent' text count: {dspa.count()}")
    
    # Check for 'Describe' in full page text
    body_text = page.inner_text('body')
    has_describe = 'Describe' in body_text
    print(f"Page body contains 'Describe': {has_describe}")
    if has_describe:
        idx = body_text.index('Describe')
        print(f"  Context: ...{body_text[max(0,idx-50):idx+80]}...")
    
    # Check for 'Suggested' 
    has_suggested = 'Suggested' in body_text
    print(f"Page body contains 'Suggested': {has_suggested}")
    
    # Check for 'Selected source'
    has_selected = 'Selected source' in body_text
    print(f"Page body contains 'Selected source': {has_selected}")

    # List all visible buttons
    buttons = page.locator('button:visible')
    bc = buttons.count()
    print(f"\nVisible buttons ({bc}):")
    for i in range(min(bc, 30)):
        try:
            txt = buttons.nth(i).inner_text(timeout=1000)
            print(f"  [{i}] {txt[:60]}")
        except:
            pass

    # Check for Shadow DOM and page content
    # Try Playwright's built-in methods
    print("\n=== Trying Playwright methods ===")
    
    # get_by_placeholder
    p1 = page.get_by_placeholder("Describe what you're looking for")
    print(f"get_by_placeholder count: {p1.count()}")
    
    # get_by_text
    p2 = page.get_by_text("Describe what you're looking for")
    print(f"get_by_text 'Describe': {p2.count()}")
    
    # get_by_role textbox
    p3 = page.get_by_role("textbox")
    print(f"get_by_role textbox: {p3.count()}")
    
    # Look for any aria-placeholder
    p4 = page.locator('[aria-placeholder]')
    print(f"aria-placeholder count: {p4.count()}")
    
    # Dump the HTML around the DSPM area
    result = page.evaluate('''() => {
        // Find the "Data Security Posture Agent" heading
        const allEls = document.querySelectorAll('*');
        for (const el of allEls) {
            if (el.textContent && el.textContent.includes('Describe what') && el.children.length < 3) {
                return {
                    tag: el.tagName,
                    className: el.className,
                    id: el.id,
                    role: el.getAttribute('role'),
                    contenteditable: el.getAttribute('contenteditable'),
                    placeholder: el.getAttribute('placeholder'),
                    ariaPlaceholder: el.getAttribute('aria-placeholder'),
                    outerHTML: el.outerHTML.substring(0, 500),
                    parentTag: el.parentElement ? el.parentElement.tagName : '',
                    parentClass: el.parentElement ? (el.parentElement.className || '').substring(0, 100) : '',
                    parentRole: el.parentElement ? el.parentElement.getAttribute('role') : '',
                    parentContenteditable: el.parentElement ? el.parentElement.getAttribute('contenteditable') : '',
                    parentOuterHTML: el.parentElement ? el.parentElement.outerHTML.substring(0, 800) : ''
                };
            }
        }
        return "Not found";
    }''')
    print(f"\n=== Element containing 'Describe what' ===")
    if isinstance(result, dict):
        for k, v in result.items():
            print(f"  {k}: {v}")
    else:
        print(f"  {result}")

    browser.close()
