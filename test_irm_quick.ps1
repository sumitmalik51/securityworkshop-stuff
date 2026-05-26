# test_irm_quick.ps1 - Quick IRM policy test on single account
param([string]$Sheet = "Batch1", [int]$Row = 0)

Import-Module ExchangeOnlineManagement -MinimumVersion 3.0

# Load creds via ImportExcel or COM
Add-Type -AssemblyName Microsoft.Office.Interop.Excel 2>$null
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open("C:\certs\spns.xlsx")
$ws = $wb.Sheets.Item($Sheet)

# Find columns
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

Write-Host "Connecting to SCC..."
try {
    Connect-IPPSSession -Credential $cred -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "Connected!"
} catch {
    Write-Host "Connection failed: $($_.Exception.Message)"
    exit 1
}

Write-Host ""
Write-Host "Getting IRM policies..."
try {
    $policies = Get-InsiderRiskPolicy -ErrorAction Stop
    Write-Host "Found $($policies.Count) policies"
    Write-Host ""
    
    foreach ($p in $policies) {
        Write-Host "=== $($p.Name) ==="
        Write-Host "  DisplayName: $($p.DisplayName)"
        Write-Host "  Enabled: $($p.IsEnabled)"
        Write-Host "  Mode: $($p.Mode)"
        Write-Host "  Scenario: $($p.InsiderRiskScenario)"
        
        # Parse indicators to check if any are enabled
        try {
            $indicators = $p.Indicators | ConvertFrom-Json
            $enabledCount = ($indicators | Where-Object { $_.Enabled -eq $true }).Count
            $totalCount = $indicators.Count
            Write-Host "  Indicators: $enabledCount enabled / $totalCount total"
            
            if ($enabledCount -eq 0) {
                Write-Host "  *** WARNING: No indicators selected! ***"
            }
        } catch {
            Write-Host "  Indicators: Could not parse"
        }
        
        # Check extensible indicators
        try {
            $extInd = $p.ExtensibleIndicators | ConvertFrom-Json
            $extEnabled = ($extInd | Where-Object { $_.Enabled -eq $true }).Count
            Write-Host "  ExtensibleIndicators: $extEnabled enabled"
        } catch {}
        
        Write-Host ""
    }
} catch {
    Write-Host "Error getting policies: $($_.Exception.Message)"
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
