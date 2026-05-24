#Requires -Modules ExchangeOnlineManagement, ImportExcel

<#
.SYNOPSIS
    Runs Purview-AutoSetup.ps1 for all tenants listed in spns.xlsx.

.DESCRIPTION
    Reads the "All" sheet from spns.xlsx (columns: AppId, AppSecret, orgname, odluser)
    and runs the automation script for each tenant sequentially using the shared certificate.

.EXAMPLE
    .\Run-AllTenants.ps1
    .\Run-AllTenants.ps1 -InputFile "C:\path\to\spns.xlsx"
#>

param(
    [string]$InputFile = "C:\Users\SumitKumar\Downloads\spns.xlsx",
    [string]$SheetName = "All",
    [string]$Thumbprint = "B1A0174A22B24D762C1C111B19C9A92D0ED488C4",
    [string]$ScriptPath = "$PSScriptRoot\Purview-AutoSetup.ps1"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $InputFile)) {
    Write-Host "ERROR: Excel file not found at $InputFile" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $ScriptPath)) {
    Write-Host "ERROR: Script not found at $ScriptPath" -ForegroundColor Red
    exit 1
}

$tenants = Import-Excel -Path $InputFile -WorksheetName $SheetName

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Purview Auto-Setup - Batch Runner" -ForegroundColor Cyan
Write-Host " Source: $InputFile (Sheet: $SheetName)" -ForegroundColor Cyan
Write-Host " Tenants to process: $($tenants.Count)" -ForegroundColor Cyan
Write-Host " Certificate: $Thumbprint" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$results = @()

foreach ($tenant in $tenants) {
    Write-Host "`n`n" -NoNewline
    Write-Host "############################################################" -ForegroundColor White
    Write-Host "# TENANT: $($tenant.orgname)" -ForegroundColor White
    Write-Host "# User:   $($tenant.odluser)" -ForegroundColor White
    Write-Host "############################################################" -ForegroundColor White

    $startTime = Get-Date

    try {
        & $ScriptPath -Thumbprint $Thumbprint -ClientId $tenant.AppId -Organization $tenant.orgname -OdlUser $tenant.odluser
        $status = "SUCCESS"
    }
    catch {
        Write-Host "  TENANT FAILED: $($_.Exception.Message)" -ForegroundColor Red
        $status = "FAILED: $($_.Exception.Message)"
    }

    $elapsed = (Get-Date) - $startTime
    $results += [PSCustomObject]@{
        Organization = $tenant.orgname
        OdlUser      = $tenant.odluser
        Status       = $status
        Duration     = $elapsed.ToString("mm\:ss")
    }
}

# Final report
Write-Host "`n`n" -NoNewline
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host " BATCH COMPLETE - All Tenants" -ForegroundColor Magenta
Write-Host "============================================================" -ForegroundColor Magenta
$results | Format-Table -AutoSize
