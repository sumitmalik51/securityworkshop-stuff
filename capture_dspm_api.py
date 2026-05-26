#!/usr/bin/env python3
"""Capture network requests made by Purview portal when loading DSPM Posture Agent tab."""

import time
import json
import sys
from playwright.sync_api import sync_playwright

USERNAME = "odl_user_2227202@otuwacne104140.onmicrosoft.com"
PASSWORD = "ojro09JCR*tW"
TENANT_ID = "509053f4-22dc-4dde-b965-5d669e8a1824"

captured = []

def on_request(request):
    url = request.url
    # Filter for API calls (skip static assets, telemetry, etc.)
    if any(skip in url for skip in [
        'browser.events.data', 'telemetry', '.js', '.css', '.png', '.svg', '.woff',
        'login.microsoftonline', 'login.live', 'msauth', 'favicon',
        'dc.services.visualstudio', 'self.events.data', 'browser.pipe',
        'config.edge', 'ntp.msn', 'bing.com', 'clarity.ms',
        'res.cdn.office.net', 'res-h2.public', 'shellprod.msocdn',
        'substrate.office', 'oteljs', 'aria.microsoft'
    ]):
        return
    method = request.method
    headers = request.headers
    post = None
    if method == "POST":
        try:
            post = request.post_data
        except:
            pass
    is_dspm = any(kw in url.lower() for kw in DSPM_KEYWORDS)
    entry = {
        "url": url,
        "method": method,
        "auth": "Bearer" if "authorization" in headers else "none",
    }
    if is_dspm:
        # Capture full headers for DSPM calls
        entry["headers"] = {k: v[:200] for k, v in headers.items() if k.lower() not in ('cookie',)}
        entry["dspm_related"] = True
    if post:
        entry["body"] = post[:500]
    captured.append(entry)
    print(f"  [{method}] {url[:150]}")


DSPM_KEYWORDS = ['dspm', 'medeina', 'DiscoveryAgent', 'discoveryagent', 'insiderrisk', 'posture', 'copilot']

def on_response(response):
    url = response.url
    if any(skip in url for skip in [
        'browser.events.data', 'telemetry', '.js', '.css', '.png', '.svg', '.woff',
        'login.microsoftonline', 'login.live', 'msauth', 'favicon',
        'dc.services.visualstudio', 'self.events.data', 'browser.pipe',
        'config.edge', 'ntp.msn', 'bing.com', 'clarity.ms',
        'res.cdn.office.net', 'res-h2.public', 'shellprod.msocdn',
        'substrate.office', 'oteljs', 'aria.microsoft'
    ]):
        return
    status = response.status
    content_type = response.headers.get("content-type", "")
    if "json" in content_type:
        try:
            body = response.text()
            is_dspm = any(kw in url.lower() for kw in DSPM_KEYWORDS)
            # For DSPM endpoints, save full response; for others, preview only
            max_preview = 50000 if is_dspm else 1000
            if len(body) > 50 or is_dspm:
                for entry in captured:
                    if entry["url"] == url and "response" not in entry:
                        entry["response_size"] = len(body)
                        entry["response_preview"] = body[:max_preview]
                        entry["status_code"] = status
                        if is_dspm:
                            entry["dspm_related"] = True
                        break
        except:
            pass


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Login
        print("=== Logging in ===")
        page.goto("https://purview.microsoft.com", wait_until="domcontentloaded", timeout=60000)
        time.sleep(5)

        try:
            page.wait_for_selector('input[type="email"]', timeout=30000)
            page.fill('input[type="email"]', USERNAME)
            page.click('input[type="submit"]')
            time.sleep(4)

            # Switch from TAP to password
            try:
                pw_link = page.locator('a:has-text("Use your password instead"), a:has-text("password"), #idA_PWD_SwitchToPassword')
                pw_link.first.click(timeout=8000)
                time.sleep(3)
            except:
                pass

            page.wait_for_selector('input[type="password"]', timeout=10000)
            page.fill('input[type="password"]', PASSWORD)
            page.click('input[type="submit"]')
            time.sleep(4)

            try:
                yes_btn = page.locator('input[type="submit"][value="Yes"], button:has-text("Yes")')
                yes_btn.first.click(timeout=15000)
            except:
                pass
            time.sleep(5)
        except Exception as e:
            print(f"Login issue: {e}")

        time.sleep(5)
        print("=== Logged in ===\n")

        # Now start capturing
        page.on("request", on_request)
        page.on("response", on_response)

        # Navigate to DSPM Asset Explorer -> Posture Agent
        dspm_url = f"https://purview.microsoft.com/dspm/assetexplorer?tid={TENANT_ID}"
        print(f"=== Navigating to DSPM: {dspm_url} ===")
        page.goto(dspm_url, wait_until="domcontentloaded", timeout=60000)
        time.sleep(10)

        # Close popups
        try:
            btn = page.locator('button:has-text("Get started")')
            if btn.count() > 0 and btn.first.is_visible():
                btn.first.click(timeout=3000)
                time.sleep(2)
        except:
            pass

        # Click Posture agent tab
        print("\n=== Clicking Posture Agent tab ===")
        try:
            tab = page.get_by_text("Posture agent").first
            tab.click(timeout=10000)
            print("  Clicked Posture agent tab")
        except Exception as e:
            print(f"  Could not click tab: {e}")

        # Wait for content to load (longer wait to catch lazy-loaded DSPM agent calls)
        print("\n=== Waiting 60s for all API calls to settle ===")
        time.sleep(60)

        # Read the Posture Agent tab DOM content to see current status
        print("\n=== Reading Posture Agent tab UI content ===")
        try:
            body_text = page.evaluate("() => document.body.innerText")
            # Extract relevant lines
            lines = body_text.split('\n')
            posture_section = False
            relevant = []
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if any(kw in line.lower() for kw in ['posture agent', 'discovery agent', 'stop generating', 
                    'running', 'rerun', 'sensitive', 'credentials', 'password', 'view insights',
                    'nothing to show', 'data sources', 'add data', 'prompt', 'get files']):
                    relevant.append(line)
            print("  Relevant UI text:")
            for r in relevant[:30]:
                print(f"    > {r}")
        except Exception as e:
            print(f"  DOM read error: {e}")

        # Print DSPM-specific calls with headers
        dspm_calls = [e for e in captured if e.get('dspm_related')]
        print(f"\n{'='*60}")
        print(f"DSPM-RELATED API CALLS ({len(dspm_calls)} of {len(captured)} total)")
        print(f"{'='*60}")

        for i, entry in enumerate(dspm_calls, 1):
            print(f"\n--- DSPM Call {i} ---")
            print(f"  Method: {entry['method']}")
            print(f"  URL: {entry['url']}")
            print(f"  Auth: {entry['auth']}")
            if 'headers' in entry:
                print(f"  Headers:")
                for k, v in entry['headers'].items():
                    if k.lower() in ('authorization', 'x-requestid', 'x-correlationid', 'content-type', 
                                     'x-ms-client-request-id', 'accept', 'x-requested-with'):
                        print(f"    {k}: {v}")
            if 'body' in entry:
                print(f"  Body: {entry['body']}")
            if 'status_code' in entry:
                print(f"  Status: {entry['status_code']}")
            if 'response_preview' in entry:
                print(f"  Response size: {entry['response_size']}")
                print(f"  Response: {entry['response_preview'][:3000]}")

        # Save to file
        with open(r"C:\certs\dspm_api_capture.json", "w") as f:
            json.dump(captured, f, indent=2)
        print(f"\nSaved {len(captured)} calls to dspm_api_capture.json")

        input("\nPress Enter to close browser...")
        browser.close()


if __name__ == "__main__":
    main()
