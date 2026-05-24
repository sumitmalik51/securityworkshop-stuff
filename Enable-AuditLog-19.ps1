param(
    [string]$InputFile = "C:\certs\66503-User-Detail-Report-separate.xlsx"
)

$ErrorActionPreference = "Continue"

$targetTenants = @(
    'otuwacne103898','otuwacne103903','otuwacne103909','otuwacne103910','otuwacne103916',
    'otuwacne103940','otuwacne103973','otuwacne103976','otuwacne103992','otuwacne104033',
    'otuwacne104035','otuwacne104090','otuwacne104098','otuwacne104102','otuwacne104108',
    'otuwacne104109','otuwacne104126','otuwacne103974','otuwacne104119'
)

# Search all relevant sheets for these tenants
$found = @{}
foreach ($ws in @('All Environments Details','Environment Detail 5','Environment Detail 6','Environment Detail 7','Environment Detail 8','Environment Detail 9')) {
    $data = Import-Excel -Path $InputFile -WorksheetName $ws
    foreach ($r in $data) {
        if ($r.username -and $r.username -match '@') {
            $dom = $r.username.Split('@')[1].Replace('.onmicrosoft.com','')
            if ($targetTenants -contains $dom -and -not $found.ContainsKey($dom)) {
                $found[$dom] = [PSCustomObject]@{User=$r.username; Pass=$r.password; Domain=$r.username.Split('@')[1]; Tenant=$dom}
            }
        }
    }
}

Write-Host "Found $($found.Count) of $($targetTenants.Count) tenants"
Write-Host ""

$results = @()
$i = 0
foreach ($tenant in $targetTenants) {
    $i++
    $info = $found[$tenant]
    if (-not $info) {
        Write-Host "[$i/19] $tenant - NOT FOUND" -ForegroundColor Red
        $results += [PSCustomObject]@{Num=$i; Tenant=$tenant; Status="NOT_FOUND"}
        continue
    }

    Write-Host "[$i/19] $($info.User) ..." -NoNewline

    $auditResult = powershell.exe -NoProfile -ExecutionPolicy Bypass -Command {
        param($user, $pass, $domain)
        $env:AZURE_ENABLE_WAM = "false"
        [System.Environment]::SetEnvironmentVariable("Broker_Enabled","false","Process")
        $sp = ConvertTo-SecureString $pass -AsPlainText -Force
        $cr = New-Object System.Management.Automation.PSCredential($user, $sp)
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -Credential $cr -ShowBanner:$false -ErrorAction Stop

        # Force enable
        Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -ErrorAction Stop
        Start-Sleep -Seconds 5

        # Verify immediately
        $config = Get-AdminAuditLogConfig
        $enabled = $config.UnifiedAuditLogIngestionEnabled

        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

        if ($enabled -eq $true) { Write-Output "CONFIRMED_ENABLED" } else { Write-Output "SET_BUT_NOT_YET_ACTIVE" }
    } -Args $info.User, $info.Pass, $info.Domain 2>&1

    $output = ($auditResult | Out-String).Trim()

    if ($output -match "CONFIRMED_ENABLED") {
        Write-Host " ENABLED + CONFIRMED" -ForegroundColor Green
        $results += [PSCustomObject]@{Num=$i; Tenant=$tenant; Status="ENABLED_CONFIRMED"}
    } elseif ($output -match "SET_BUT_NOT_YET_ACTIVE") {
        Write-Host " SET (pending propagation)" -ForegroundColor Yellow
        $results += [PSCustomObject]@{Num=$i; Tenant=$tenant; Status="SET_PENDING"}
    } else {
        Write-Host " ERROR" -ForegroundColor Red
        Write-Host "  $output"
        $results += [PSCustomObject]@{Num=$i; Tenant=$tenant; Status="ERROR"}
    }
}

Write-Host ""
Write-Host "================================="
Write-Host "ENABLE SUMMARY"
Write-Host "================================="
$results | Format-Table -AutoSize
$confirmed = ($results | Where-Object {$_.Status -eq 'ENABLED_CONFIRMED'}).Count
$pending = ($results | Where-Object {$_.Status -eq 'SET_PENDING'}).Count
$errors = ($results | Where-Object {$_.Status -like 'ERROR*'}).Count
Write-Host "Confirmed: $confirmed | Pending: $pending | Errors: $errors"
