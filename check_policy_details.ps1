# check_policy_details.ps1 - Get detailed IRM policy info
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
Connect-IPPSSession -Credential $cred -ErrorAction Stop -WarningAction SilentlyContinue

$policies = Get-InsiderRiskPolicy -ErrorAction Stop
foreach ($p in $policies) {
    if ($p.InsiderRiskScenario -eq "TenantSetting") { continue }
    Write-Host ""
    Write-Host "=== $($p.Name) ==="
    Write-Host "  Scenario: $($p.InsiderRiskScenario)"
    Write-Host "  Mode: $($p.Mode)"
    Write-Host "  Enabled: $($p.Enabled)"
    Write-Host "  DistributionStatus: $($p.DistributionStatus)"
    Write-Host "  DistributionSyncStatus: $($p.DistributionSyncStatus)"
    Write-Host "  CreatedBy: $($p.CreatedBy)"
    Write-Host "  CreatedUTC: $($p.CreationTimeUtc)"
    Write-Host "  ModifiedUTC: $($p.ModificationTimeUtc)"
    
    try {
        $ph = $p.PolicyHealth | ConvertFrom-Json
        Write-Host "  PolicyHealth: $($ph.HealthStatus) | Unhealthy: $($ph.UnhealthyCount) | Recs: $($ph.RecommendReviewCount)"
        if ($ph.ValidationDetails -and $ph.ValidationDetails.Count -gt 0) {
            foreach ($vd in $ph.ValidationDetails) {
                Write-Host "    ValidationDetail: $($vd.Name) - $($vd.ValidationSeverity)"
            }
        }
    } catch {
        Write-Host "  PolicyHealth: Could not parse"
    }
    
    Write-Host "  ExchangeLocation: $($p.ExchangeLocation)"
    Write-Host "  Workload: $($p.Workload)"
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
