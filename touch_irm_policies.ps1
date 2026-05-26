# touch_irm_policies.ps1 - Force-update IRM policies to trigger redistribution
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

$policies = Get-InsiderRiskPolicy -ErrorAction SilentlyContinue
Write-Host "`nBefore update:"
foreach ($p in $policies) {
    if ($p.InsiderRiskScenario -eq "TenantSetting") { continue }
    Write-Host "  $($p.Name) | DistStatus: $($p.DistributionStatus) | Enabled: $($p.Enabled) | Mode: $($p.Mode)"
}

Write-Host "`n=== Forcing policy updates ==="
foreach ($p in $policies) {
    if ($p.InsiderRiskScenario -eq "TenantSetting") { continue }
    Write-Host "  Updating: $($p.Name)..."
    try {
        # Try to re-enable / force mode to trigger redistribution
        Set-InsiderRiskPolicy -Identity $p.Name -Enabled $true -ErrorAction Stop
        Write-Host "    Updated (re-enabled)!"
    } catch {
        Write-Host "    Set error: $($_.Exception.Message)"
        # Try RetryDistribution if available
        try {
            Set-InsiderRiskPolicy -Identity $p.Name -RetryDistribution -ErrorAction Stop
            Write-Host "    RetryDistribution succeeded!"
        } catch {
            Write-Host "    RetryDistribution error: $($_.Exception.Message)"
        }
    }
}

Write-Host "`nWaiting 10 seconds..."
Start-Sleep -Seconds 10

Write-Host "`nAfter update:"
$policies2 = Get-InsiderRiskPolicy -ErrorAction SilentlyContinue
foreach ($p in $policies2) {
    if ($p.InsiderRiskScenario -eq "TenantSetting") { continue }
    Write-Host "  $($p.Name) | DistStatus: $($p.DistributionStatus) | DistSync: $($p.DistributionSyncStatus) | Modified: $($p.ModificationTimeUtc)"
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
Write-Host "`nDone. Refresh the Purview UI to check if policies now appear."
