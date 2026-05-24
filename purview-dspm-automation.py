"""
Purview DSPM Automation
  - Exercise 3 Task 2: Run natural-language discovery in DSPM
  - Exercise 4 Task 1: Create Credential Scanning Task in DSI

Runs across all tenants from Excel file using Playwright browser automation.
"""

import asyncio
import json
import os
import sys
from datetime import datetime

try:
    from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout
except ImportError:
    os.system(f"{sys.executable} -m pip install playwright openpyxl")
    os.system(f"{sys.executable} -m playwright install chromium")
    from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

try:
    from openpyxl import load_workbook
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl")
    from openpyxl import load_workbook


# ─── Configuration ───────────────────────────────────────────────
INPUT_FILE = r"C:\certs\66503-User-Detail-Report-separate.xlsx"
WORKSHEETS = [
    "Environment Detail 5",
    "Environment Detail 6",
    "Environment Detail 7",
    "Environment Detail 8",
    "Environment Detail 9",
]
PURVIEW_URL = "https://purview.microsoft.com"
TIMEOUT = 60000          # 60s for navigation/waits
LONG_TIMEOUT = 120000    # 120s for job submission
HEADLESS = False         # Set True for unattended runs
LOG_FILE = r"C:\certs\purview_dspm_results.json"

# Exercise 3 Task 2 config
DSPM_PROMPT = "Find sensitive data in my selected data sources."

# Exercise 4 Task 1 config
CRED_SCAN_TASK_NAME = "Weekly Engineering Credential Exposure Scan"
CRED_SCAN_CONTEXT = (
    "Perform a credential exposure investigation across the Engineering "
    "collaboration OneDrive scope. Identify files containing high-risk secrets "
    "such as GitHub personal access tokens, GitLab tokens, Docker credentials, "
    "cloud access secrets, and API keys. Prioritize active or production-related "
    "credentials that may create unauthorized access risk. Deprioritize sample "
    "credentials, placeholder values, or expired development tokens unless they "
    "appear alongside sensitive project content."
)


def load_tenants(file_path, worksheets):
    """Load tenant credentials from Excel sheets."""
    wb = load_workbook(file_path, read_only=True, data_only=True)
    tenants = []
    for ws_name in worksheets:
        if ws_name not in wb.sheetnames:
            print(f"  WARNING: Sheet '{ws_name}' not found, skipping")
            continue
        ws = wb[ws_name]
        headers = [str(c.value).strip().lower() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
        user_col = headers.index("username") if "username" in headers else None
        pass_col = headers.index("password") if "password" in headers else None
        if user_col is None or pass_col is None:
            print(f"  WARNING: Sheet '{ws_name}' missing username/password columns")
            continue
        for row in ws.iter_rows(min_row=2, values_only=True):
            user = row[user_col]
            pwd = row[pass_col]
            if user and pwd:
                tenants.append({
                    "username": str(user).strip(),
                    "password": str(pwd).strip(),
                    "sheet": ws_name,
                    "domain": str(user).strip().split("@")[1] if "@" in str(user) else "",
                })
    wb.close()
    print(f"Loaded {len(tenants)} tenants from {len(worksheets)} sheets")
    return tenants


async def login_purview(page, username, password):
    """Login to Microsoft Purview portal."""
    await page.goto(PURVIEW_URL, wait_until="domcontentloaded", timeout=TIMEOUT)
    await page.wait_for_timeout(3000)

    # Enter email
    email_input = page.locator('input[type="email"], input[name="loginfmt"]')
    await email_input.wait_for(state="visible", timeout=TIMEOUT)
    await email_input.fill(username)
    await page.locator('input[type="submit"], #idSIButton9').click()
    await page.wait_for_timeout(2000)

    # Enter password
    pwd_input = page.locator('input[type="password"], input[name="passwd"]')
    await pwd_input.wait_for(state="visible", timeout=TIMEOUT)
    await pwd_input.fill(password)
    await page.locator('input[type="submit"], #idSIButton9').click()
    await page.wait_for_timeout(3000)

    # "Stay signed in?" - click No
    try:
        stay_btn = page.locator('#idBtn_Back, text="No"')
        await stay_btn.wait_for(state="visible", timeout=10000)
        await stay_btn.click()
    except PlaywrightTimeout:
        pass

    # Wait for Purview portal to load
    await page.wait_for_timeout(5000)
    await page.wait_for_load_state("networkidle", timeout=TIMEOUT)
    print("    Logged in to Purview portal")


async def run_dspm_discovery(page, username):
    """
    Exercise 3 Task 2: Run natural-language discovery in DSPM.
    
    Flow:
    1. Open DSPM -> Discover -> Asset explorer
    2. Select Posture Agent (preview)
    3. Add data sources -> search ODL_User_, select, Sites only, Save
    4. Enter prompt: "Find sensitive data in my selected data sources."
    5. Click Submit -> Confirm
    """
    result = {"task": "DSPM Discovery", "status": "unknown", "details": ""}
    odl_user = username.split("@")[0] if "@" in username else username

    try:
        # Navigate to DSPM - try left nav
        try:
            dspm_link = page.get_by_text("Data Security Posture Management").first
            await dspm_link.wait_for(state="visible", timeout=10000)
            await dspm_link.click()
        except PlaywrightTimeout:
            # Try Solutions menu or direct nav
            try:
                solutions = page.get_by_text("Solutions").first
                await solutions.click()
                await page.wait_for_timeout(2000)
                dspm = page.get_by_text("Data Security Posture Management").first
                await dspm.click()
            except PlaywrightTimeout:
                # Try DSPM direct URL
                await page.goto(
                    "https://purview.microsoft.com/dataseurityposturemanagement",
                    wait_until="domcontentloaded", timeout=TIMEOUT
                )
        await page.wait_for_timeout(3000)

        # Navigate to Discover -> Asset explorer
        try:
            discover_link = page.get_by_text("Discover").first
            await discover_link.wait_for(state="visible", timeout=15000)
            await discover_link.click()
            await page.wait_for_timeout(2000)
        except PlaywrightTimeout:
            pass

        asset_explorer = page.get_by_text("Asset explorer").first
        await asset_explorer.wait_for(state="visible", timeout=15000)
        await asset_explorer.click()
        await page.wait_for_timeout(3000)

        # Select "Posture Agent (preview)"
        posture_tab = page.get_by_text("Posture Agent").first
        await posture_tab.wait_for(state="visible", timeout=15000)
        await posture_tab.click()
        await page.wait_for_timeout(3000)

        # Click "Add data sources"
        add_btn = page.get_by_role("button", name="Add data sources").or_(
            page.get_by_text("Add data sources").first
        )
        await add_btn.wait_for(state="visible", timeout=15000)
        await add_btn.click()
        await page.wait_for_timeout(3000)

        # Search for ODL_User_
        search_input = page.locator(
            'input[type="search"], input[placeholder*="Search"], '
            'input[placeholder*="search"], input[aria-label*="Search"]'
        ).first
        await search_input.wait_for(state="visible", timeout=15000)
        await search_input.fill(odl_user)
        await page.wait_for_timeout(3000)

        # Select the user from results
        try:
            user_result = page.locator(f'text="{odl_user}"').nth(1)
            await user_result.wait_for(state="visible", timeout=10000)
            await user_result.click()
        except PlaywrightTimeout:
            first_result = page.locator(
                '[role="option"], [role="row"], [role="listitem"]'
            ).first
            await first_result.click()
        await page.wait_for_timeout(1000)

        # Ensure Sites is selected, Mailbox unchecked
        try:
            mailbox_cb = page.get_by_label("Mailbox")
            if await mailbox_cb.is_visible():
                if await mailbox_cb.is_checked():
                    await mailbox_cb.uncheck()
        except Exception:
            pass

        # Click Save
        save_btn = page.get_by_role("button", name="Save").first
        await save_btn.wait_for(state="visible", timeout=10000)
        await save_btn.click()
        await page.wait_for_timeout(3000)

        # Enter prompt
        prompt_input = page.locator(
            'textarea, input[placeholder*="prompt"], '
            'input[placeholder*="Enter"], [contenteditable="true"]'
        ).first
        await prompt_input.wait_for(state="visible", timeout=15000)
        await prompt_input.fill(DSPM_PROMPT)
        await page.wait_for_timeout(1000)

        # Click Submit
        submit_btn = page.get_by_role("button", name="Submit").or_(
            page.get_by_text("Submit").first
        )
        await submit_btn.wait_for(state="visible", timeout=10000)
        await submit_btn.click()
        await page.wait_for_timeout(5000)

        # Click Confirm
        try:
            confirm_btn = page.get_by_role("button", name="Confirm").first
            await confirm_btn.wait_for(state="visible", timeout=15000)
            await confirm_btn.click()
            await page.wait_for_timeout(5000)
        except PlaywrightTimeout:
            pass

        # Verify job submitted
        try:
            job_status = page.locator(
                'text="Job estimation in progress"'
            ).or_(page.locator('text="in progress"')
            ).or_(page.locator('text="Queued"')
            ).or_(page.locator('text="Running"'))
            await job_status.wait_for(state="visible", timeout=15000)
            result["status"] = "submitted"
            result["details"] = f"Discovery job submitted for {odl_user}"
            print(f"    DSPM Discovery: Job submitted")
        except PlaywrightTimeout:
            result["status"] = "submitted_unverified"
            result["details"] = f"Submitted (status unverified) for {odl_user}"
            print(f"    DSPM Discovery: Submitted (unverified)")

    except PlaywrightTimeout as e:
        result["status"] = "error"
        result["details"] = f"Timeout: {str(e)[:200]}"
        print(f"    DSPM Discovery: ERROR - {str(e)[:100]}")
    except Exception as e:
        result["status"] = "error"
        result["details"] = f"Error: {str(e)[:200]}"
        print(f"    DSPM Discovery: ERROR - {str(e)[:100]}")

    return result


async def create_credential_scan(page, username):
    """
    Exercise 4 Task 1: Create Credential Scanning Task in DSI.
    
    Flow:
    1. Navigate to Agents -> Posture Agent -> Open in -> DSI
    2. + New task -> Credential scanning
    3. Enter task name, configure data sources, enter AI context
    4. Click Create
    """
    result = {"task": "Credential Scanning", "status": "unknown", "details": ""}
    odl_user = username.split("@")[0] if "@" in username else username

    try:
        # Navigate to Purview home
        await page.goto(PURVIEW_URL, wait_until="domcontentloaded", timeout=TIMEOUT)
        await page.wait_for_timeout(5000)

        # Click "Agents"
        agents_link = page.get_by_role("link", name="Agents").or_(
            page.get_by_text("Agents").first
        )
        await agents_link.wait_for(state="visible", timeout=TIMEOUT)
        await agents_link.click()
        await page.wait_for_timeout(3000)

        # Click "Explore agents"
        explore_btn = page.get_by_text("Explore agents").first
        await explore_btn.wait_for(state="visible", timeout=15000)
        await explore_btn.click()
        await page.wait_for_timeout(3000)

        # Posture Agent -> "Open in" -> "Data Security Investigations"
        open_in_btn = page.get_by_text("Open in").first
        await open_in_btn.wait_for(state="visible", timeout=15000)
        await open_in_btn.click()
        await page.wait_for_timeout(2000)

        dsi_option = page.get_by_text("Data Security Investigations").first
        await dsi_option.wait_for(state="visible", timeout=10000)
        await dsi_option.click()
        await page.wait_for_timeout(5000)
        await page.wait_for_load_state("networkidle", timeout=TIMEOUT)

        # Click "+ New task"
        new_task_btn = page.get_by_role("button", name="New task").or_(
            page.get_by_text("New task").first
        )
        await new_task_btn.wait_for(state="visible", timeout=15000)
        await new_task_btn.click()
        await page.wait_for_timeout(2000)

        # Select "Credential scanning"
        cred_option = page.get_by_text("Credential scanning").first
        await cred_option.wait_for(state="visible", timeout=10000)
        await cred_option.click()
        await page.wait_for_timeout(3000)

        # Enter Task Name
        task_name_input = page.locator(
            'input[placeholder*="name"], input[placeholder*="Name"], '
            'input[aria-label*="Task name"], input[aria-label*="Name"]'
        ).first
        await task_name_input.wait_for(state="visible", timeout=15000)
        await task_name_input.fill(CRED_SCAN_TASK_NAME)
        await page.wait_for_timeout(1000)

        # Click "Edit" on Data sources
        edit_btn = page.get_by_role("button", name="Edit").or_(
            page.get_by_text("Edit").first
        )
        await edit_btn.wait_for(state="visible", timeout=10000)
        await edit_btn.click()
        await page.wait_for_timeout(3000)

        # Remove default "All people and groups"
        try:
            remove_btns = page.locator(
                '[aria-label*="Remove"], [aria-label*="remove"], '
                'button:has-text("Remove"), .ms-TagItem-close, '
                '[data-icon-name="Cancel"], [aria-label*="Close"]'
            )
            count = await remove_btns.count()
            for _ in range(count):
                await remove_btns.first.click()
                await page.wait_for_timeout(500)
        except Exception:
            pass

        # Search and select ODL_User_
        search_input = page.locator(
            'input[type="search"], input[placeholder*="Search"], '
            'input[placeholder*="search"], input[aria-label*="Search"]'
        ).first
        await search_input.wait_for(state="visible", timeout=10000)
        await search_input.fill(odl_user)
        await page.wait_for_timeout(3000)

        # Select user from results
        try:
            user_result = page.locator(f'text="{odl_user}"').nth(1)
            await user_result.wait_for(state="visible", timeout=10000)
            await user_result.click()
        except PlaywrightTimeout:
            first_result = page.locator(
                '[role="option"], [role="row"], [role="listitem"]'
            ).first
            await first_result.click()
        await page.wait_for_timeout(1000)

        # Ensure Sites only (uncheck Mailbox)
        try:
            mailbox_cb = page.get_by_label("Mailbox")
            if await mailbox_cb.is_visible() and await mailbox_cb.is_checked():
                await mailbox_cb.uncheck()
        except Exception:
            pass

        # Click Save
        save_btn = page.get_by_role("button", name="Save").first
        await save_btn.wait_for(state="visible", timeout=10000)
        await save_btn.click()
        await page.wait_for_timeout(3000)

        # Enter Additional context for AI
        context_input = page.locator(
            'textarea, [aria-label*="context"], [aria-label*="Context"], '
            '[aria-label*="Additional"], [placeholder*="context"]'
        ).first
        await context_input.wait_for(state="visible", timeout=10000)
        await context_input.fill(CRED_SCAN_CONTEXT)
        await page.wait_for_timeout(1000)

        # Click Create
        create_btn = page.get_by_role("button", name="Create").first
        await create_btn.wait_for(state="visible", timeout=10000)
        await create_btn.click()
        await page.wait_for_timeout(5000)

        result["status"] = "created"
        result["details"] = f"Task created for {odl_user}"
        print(f"    Credential Scan: Task created")

    except PlaywrightTimeout as e:
        result["status"] = "error"
        result["details"] = f"Timeout: {str(e)[:200]}"
        print(f"    Credential Scan: ERROR - {str(e)[:100]}")
    except Exception as e:
        result["status"] = "error"
        result["details"] = f"Error: {str(e)[:200]}"
        print(f"    Credential Scan: ERROR - {str(e)[:100]}")

    return result


async def process_tenant(playwright, tenant, index, total):
    """Process one tenant: login -> DSPM discovery -> credential scan."""
    username = tenant["username"]
    password = tenant["password"]

    print(f"\n[{index}/{total}] {username}")
    print(f"  Sheet: {tenant['sheet']}")

    tenant_result = {
        "username": username,
        "domain": tenant["domain"],
        "sheet": tenant["sheet"],
        "timestamp": datetime.now().isoformat(),
        "dspm_discovery": None,
        "credential_scan": None,
    }

    browser = await playwright.chromium.launch(headless=HEADLESS)
    context = await browser.new_context(
        viewport={"width": 1920, "height": 1080},
        ignore_https_errors=True,
    )
    page = await context.new_page()
    page.set_default_timeout(TIMEOUT)

    try:
        await login_purview(page, username, password)
        tenant_result["dspm_discovery"] = await run_dspm_discovery(page, username)
        tenant_result["credential_scan"] = await create_credential_scan(page, username)
    except Exception as e:
        print(f"  FATAL: {str(e)[:200]}")
        tenant_result["dspm_discovery"] = tenant_result["dspm_discovery"] or {
            "task": "DSPM Discovery", "status": "fatal_error", "details": str(e)[:200]
        }
        tenant_result["credential_scan"] = tenant_result["credential_scan"] or {
            "task": "Credential Scanning", "status": "fatal_error", "details": str(e)[:200]
        }
    finally:
        await context.close()
        await browser.close()

    return tenant_result


async def main():
    print("=" * 60)
    print("Purview DSPM + DSI Automation")
    print("  Exercise 3 Task 2: DSPM Natural-Language Discovery")
    print("  Exercise 4 Task 1: DSI Credential Scanning Task")
    print("=" * 60)

    tenants = load_tenants(INPUT_FILE, WORKSHEETS)
    if not tenants:
        print("No tenants found.")
        return

    results = []
    total = len(tenants)

    async with async_playwright() as pw:
        for i, tenant in enumerate(tenants, 1):
            result = await process_tenant(pw, tenant, i, total)
            results.append(result)
            with open(LOG_FILE, "w", encoding="utf-8") as f:
                json.dump(results, f, indent=2, default=str)

    # Summary
    print("\n" + "=" * 60)
    print("FINAL SUMMARY")
    print("=" * 60)
    dspm_ok = sum(1 for r in results if r["dspm_discovery"]["status"] in ("submitted", "submitted_unverified"))
    dspm_err = sum(1 for r in results if r["dspm_discovery"]["status"] in ("error", "fatal_error"))
    cs_ok = sum(1 for r in results if r["credential_scan"]["status"] == "created")
    cs_err = sum(1 for r in results if r["credential_scan"]["status"] in ("error", "fatal_error"))

    print(f"\nDSPM Discovery:      {dspm_ok} submitted / {dspm_err} errors")
    print(f"Credential Scanning: {cs_ok} created / {cs_err} errors")
    print(f"\nResults: {LOG_FILE}")


if __name__ == "__main__":
    asyncio.run(main())
