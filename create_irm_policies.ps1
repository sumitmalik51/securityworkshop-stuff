# create_irm_policies.ps1 - Create IRM policies for a specific user's tenant
param(
    [string]$Sheet = "Batch1",
    [int]$Row = 0
)

Import-Module ExchangeOnlineManagement -MinimumVersion 3.0

# Load creds via COM
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open("C:\certs\spns.xlsx")
$ws = $wb.Sheets.Item($Sheet)

$colMap = @{}
$c = 1
while ($ws.Cells.Item(1, $c).Value2) {
    $colMap[$ws.Cells.Item(1, $c).Value2] = $c
    $c++
}

$dataRow = $Row + 2
$username = $ws.Cells.Item($dataRow, $colMap["odluser"]).Value2
$password = $ws.Cells.Item($dataRow, $colMap["odlpassword"]).Value2
$tenantId = $ws.Cells.Item($dataRow, $colMap["TenantId"]).Value2
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host "User: $username"
Write-Host "TenantID: $tenantId"
Write-Host ""

$secPwd = ConvertTo-SecureString $password -AsPlainText -Force
$cred = New-Object PSCredential($username, $secPwd)

# Disconnect any existing sessions first
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 2>$null } catch {}

Write-Host "Connecting to SCC..."
try {
    Connect-IPPSSession -Credential $cred -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "Connected!"
} catch {
    Write-Host "Connection failed: $($_.Exception.Message)"
    exit 1
}

# Check existing policies
Write-Host "`nChecking existing IRM policies..."
$existing = Get-InsiderRiskPolicy -ErrorAction SilentlyContinue
$userPolicies = $existing | Where-Object { $_.InsiderRiskScenario -ne "TenantSetting" }
Write-Host "Found $($userPolicies.Count) user policies"

if ($userPolicies.Count -gt 0) {
    Write-Host "Policies already exist:"
    $userPolicies | ForEach-Object { Write-Host "  - $($_.Name)" }
    Write-Host "`nSkipping creation."
} else {
    Write-Host "`nCreating IRM policies..."

    # Policy 1: Lab - IRM Data Leak Monitoring (LeakOfInformation with AggregatedDataLeakTrigger)
    Write-Host "`n[1/3] Creating 'Lab - IRM Data Leak Monitoring'..."
    try {
        New-InsiderRiskPolicy -Name "Lab - IRM Data Leak Monitoring" `
            -InsiderRiskScenario "LeakOfInformation" `
            -ErrorAction Stop
        Write-Host "  Created successfully!"
    } catch {
        Write-Host "  Error: $($_.Exception.Message)"
    }

    # Policy 2: Lab - IRM Data Theft Policy (IntellectualPropertyTheft with AadLeaver trigger)
    Write-Host "`n[2/3] Creating 'Lab - IRM Data Theft Policy'..."
    try {
        New-InsiderRiskPolicy -Name "Lab - IRM Data Theft Policy" `
            -InsiderRiskScenario "IntellectualPropertyTheft" `
            -ErrorAction Stop
        Write-Host "  Created successfully!"
    } catch {
        Write-Host "  Error: $($_.Exception.Message)"
    }

    # Policy 3: Lab - IRM Critical Data Leak (LeakOfInformation)
    Write-Host "`n[3/3] Creating 'Lab - IRM Critical Data Leak'..."
    try {
        New-InsiderRiskPolicy -Name "Lab - IRM Critical Data Leak" `
            -InsiderRiskScenario "LeakOfInformation" `
            -ErrorAction Stop
        Write-Host "  Created successfully!"
    } catch {
        Write-Host "  Error: $($_.Exception.Message)"
    }

    # Verify
    Write-Host "`nVerifying..."
    $verify = Get-InsiderRiskPolicy -ErrorAction SilentlyContinue
    $verifyUser = $verify | Where-Object { $_.InsiderRiskScenario -ne "TenantSetting" }
    Write-Host "Policies after creation: $($verifyUser.Count)"
    $verifyUser | ForEach-Object {
        $ph = $_.PolicyHealth | ConvertFrom-Json -ErrorAction SilentlyContinue
        Write-Host "  - $($_.Name) | Health: $($ph.HealthStatus) | Mode: $($_.Mode)"
    }
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
Write-Host "`nDone."
