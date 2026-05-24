<#
.SYNOPSIS
    Checks and enables Unified Audit Log Ingestion across tenants.
    Logs status to Excel: whether it was already enabled or manually enabled by this script.

.NOTES
    Uses credential-based EXO login (same pattern as phishing script).
    Per Microsoft docs, Get-AdminAuditLogConfig MUST run via Exchange Online PowerShell
    (NOT IPPS) — IPPS always returns False even when enabled.
#>

param(
    [string]$InputFile    = "C:\certs\66503-User-Detail-Report-separate.xlsx",
    [string]$Worksheet    = "All Environments Details"
)

$ErrorActionPreference = "Continue"

# =========================================================
# IMPORT EXCEL
# =========================================================
$users = Import-Excel -Path $InputFile -WorksheetName $Worksheet
Write-Host "Loaded $($users.Count) tenants from '$Worksheet'"
Write-Host ""

foreach ($row in $users) {

    $user     = $row.username
    $pass     = $row.password
    $tenantId = $row.'tenant id'
    $domain   = $user.Split("@")[1]

    Write-Host "================================="
    Write-Host "Processing: $user"
    Write-Host "Tenant: $domain"
    Write-Host "================================="

    # Run in subprocess to avoid DLL/broker conflicts (same pattern as phishing script)
    $auditResult = powershell.exe -NoProfile -ExecutionPolicy Bypass -Command {
        param($user, $pass, $domain)

        $env:AZURE_ENABLE_WAM = "false"
        [System.Environment]::SetEnvironmentVariable("Broker_Enabled","false","Process")

        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber
        }

        $sp = ConvertTo-SecureString $pass -AsPlainText -Force
        $cr = New-Object System.Management.Automation.PSCredential($user, $sp)

        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -Credential $cr -ShowBanner:$false -ErrorAction Stop

        # Check current status
        $config = Get-AdminAuditLogConfig
        $enabled = $config.UnifiedAuditLogIngestionEnabled

        if ($enabled -eq $true) {
            Write-Host "ALREADY_ENABLED"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            return "ALREADY_ENABLED"
        }
        else {
            # Enable it
            Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
            Write-Host "MANUALLY_ENABLED"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            return "MANUALLY_ENABLED"
        }

    } -Args $user, $pass, $domain 2>&1

    # Parse result from subprocess output
    $output = ($auditResult | Out-String).Trim()

    if ($output -match "ALREADY_ENABLED") {
        $row | Add-Member -NotePropertyName "enabled" -NotePropertyValue "Yes" -Force
        $row | Add-Member -NotePropertyName "manuallydone" -NotePropertyValue "" -Force
        Write-Host "  Status: Already Enabled" -ForegroundColor Green
    }
    elseif ($output -match "MANUALLY_ENABLED") {
        $row | Add-Member -NotePropertyName "enabled" -NotePropertyValue "No" -Force
        $row | Add-Member -NotePropertyName "manuallydone" -NotePropertyValue "Done" -Force
        Write-Host "  Status: Manually Enabled Done" -ForegroundColor Yellow
    }
    else {
        $row | Add-Member -NotePropertyName "enabled" -NotePropertyValue "Error" -Force
        $row | Add-Member -NotePropertyName "manuallydone" -NotePropertyValue "Error" -Force
        Write-Host "  Status: ERROR" -ForegroundColor Red
        Write-Host "  $output"
    }

    Write-Host ""
}

# =========================================================
# WRITE BACK TO SAME EXCEL
# =========================================================
$users | Select-Object username, password, 'tenant id', enabled, manuallydone | Export-Excel `
    -Path $InputFile `
    -WorksheetName $Worksheet `
    -AutoSize `
    -BoldTopRow `
    -FreezeTopRow `
    -ClearSheet

Write-Host "================================="
Write-Host "COMPLETE"
Write-Host "================================="
$already  = ($users | Where-Object { $_.enabled -eq "Yes" }).Count
$manual   = ($users | Where-Object { $_.manuallydone -eq "Done" }).Count
$errCount = ($users | Where-Object { $_.enabled -eq "Error" }).Count
Write-Host "Already Enabled: $already"
Write-Host "Manually Done:   $manual"
Write-Host "Errors:          $errCount"
Write-Host "Updated:         $InputFile -> '$Worksheet'"
