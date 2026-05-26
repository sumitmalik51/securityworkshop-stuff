# fix_irm_policies.ps1 - Delete stuck policies and recreate them
param([string]$Sheet = "Batch1", [int]$Row = 0)

Import-Module ExchangeOnlineManagement -MinimumVersion 3.0

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open("C:\certs\spns.xlsx")
$ws = $wb.Sheets.Item($Sheet)
$colMap = @{}; $c = 1
while ($ws.Cells.Item(1, $c).Value2) { $colMap[$ws.Cells.Item(1, $c).Value2] = $c; $c++ }
$dataRow = $Row + 2
$username = $ws.Cells.Item($dataRow, $colMap["odluser"]).Value2
$password = $ws.Cells.Item($dataRow, $colMap["odlpassword"]).Value2
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host "User: $username"
$secPwd = ConvertTo-SecureString $password -AsPlainText -Force
$cred = New-Object PSCredential($username, $secPwd)

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 2>$null } catch {}
Write-Host "Connecting..."
Connect-IPPSSession -Credential $cred -ErrorAction Stop -WarningAction SilentlyContinue

# Step 1: Delete existing stuck policies
Write-Host "`n=== Step 1: Deleting stuck policies ==="
$policies = Get-InsiderRiskPolicy -ErrorAction SilentlyContinue
foreach ($p in $policies) {
    if ($p.InsiderRiskScenario -eq "TenantSetting") { continue }
    Write-Host "  Deleting: $($p.Name) (DistStatus: $($p.DistributionStatus))..."
    try {
        Remove-InsiderRiskPolicy -Identity $p.Name -Confirm:$false -ErrorAction Stop
        Write-Host "    Deleted."
    } catch {
        Write-Host "    Error: $($_.Exception.Message)"
    }
}

Write-Host "`nWaiting 15 seconds for deletion to propagate..."
Start-Sleep -Seconds 15

# Step 2: Verify deletion
Write-Host "`n=== Step 2: Verify deletion ==="
$remaining = Get-InsiderRiskPolicy -ErrorAction SilentlyContinue
$userRemaining = $remaining | Where-Object { $_.InsiderRiskScenario -ne "TenantSetting" }
Write-Host "  Remaining user policies: $($userRemaining.Count)"

# Step 3: Recreate policies
Write-Host "`n=== Step 3: Creating fresh policies ==="

Write-Host "  [1/3] Lab - IRM Critical Data Leak..."
try {
    New-InsiderRiskPolicy -Name "Lab - IRM Critical Data Leak" -InsiderRiskScenario "LeakOfInformation" -ErrorAction Stop
    Write-Host "    Created!"
} catch { Write-Host "    Error: $($_.Exception.Message)" }

Write-Host "  [2/3] Lab - IRM Data Leak Monitoring..."
try {
    New-InsiderRiskPolicy -Name "Lab - IRM Data Leak Monitoring" -InsiderRiskScenario "LeakOfInformation" -ErrorAction Stop
    Write-Host "    Created!"
} catch { Write-Host "    Error: $($_.Exception.Message)" }

Write-Host "  [3/3] Lab - IRM Data Theft Policy..."
try {
    New-InsiderRiskPolicy -Name "Lab - IRM Data Theft Policy" -InsiderRiskScenario "IntellectualPropertyTheft" -ErrorAction Stop
    Write-Host "    Created!"
} catch { Write-Host "    Error: $($_.Exception.Message)" }

# Step 4: Verify
Write-Host "`n=== Step 4: Verify new policies ==="
Start-Sleep -Seconds 5
$final = Get-InsiderRiskPolicy -ErrorAction SilentlyContinue
foreach ($p in $final) {
    if ($p.InsiderRiskScenario -eq "TenantSetting") { continue }
    try {
        $ph = $p.PolicyHealth | ConvertFrom-Json
        $health = $ph.HealthStatus
    } catch { $health = "?" }
    Write-Host "  $($p.Name) | Enabled: $($p.Enabled) | DistStatus: $($p.DistributionStatus) | Health: $health | Created: $($p.CreationTimeUtc)"
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
Write-Host "`nDone."
