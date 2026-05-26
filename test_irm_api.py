import requests, json, openpyxl

wb = openpyxl.load_workbook('spns.xlsx', read_only=True)
ws = wb['Batch1']
headers_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
user_col = headers_row.index('odluser')
pwd_col = headers_row.index('odlpassword')
tid_col = headers_row.index('TenantId')
row = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))[0]
username = str(row[user_col]).strip()
password = str(row[pwd_col]).strip()
tenant_id = str(row[tid_col]).strip()
wb.close()
print(f"User: {username}, TID: {tenant_id}")

# Get Graph token
token_resp = requests.post(
    f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
    data={
        "grant_type": "password",
        "client_id": "d3590ed6-52b3-4102-aeff-aad2292ab01c",
        "scope": "https://graph.microsoft.com/.default",
        "username": username,
        "password": password,
    }, timeout=30
).json()
token = token_resp["access_token"]
print("Token acquired")

# Try compliance API endpoints
urls = [
    f"https://compliance.microsoft.com/api/InsiderRiskManagement/Policies?tenantId={tenant_id}",
    "https://compliance.microsoft.com/api/compliancemanager/policies",
]
for url in urls:
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    short = url.split("/api/")[1][:60]
    print(f"\n--- {short} ---")
    print(f"Status: {resp.status_code}")
    ct = resp.headers.get("content-type", "")
    print(f"Content-Type: {ct}")
    print(f"Body (500 chars): {resp.text[:500]}")

# Try with Compliance scope token
print("\n\n=== Trying with compliance scope ===")
for scope in [
    "https://compliance.microsoft.com/.default",
    "https://outlook.office365.com/.default",
]:
    token_resp2 = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "grant_type": "password",
            "client_id": "d3590ed6-52b3-4102-aeff-aad2292ab01c",
            "scope": scope,
            "username": username,
            "password": password,
        }, timeout=30
    ).json()
    if "access_token" in token_resp2:
        print(f"\n[{scope}] Token acquired")
        t2 = token_resp2["access_token"]
        for url in urls:
            resp = requests.get(url, headers={"Authorization": f"Bearer {t2}"}, timeout=30)
            short = url.split("/api/")[1][:60]
            print(f"  {short}: HTTP {resp.status_code} | {resp.text[:300]}")
    else:
        print(f"\n[{scope}] Token FAILED: {token_resp2.get('error', '')} - {token_resp2.get('error_description', '')[:150]}")
