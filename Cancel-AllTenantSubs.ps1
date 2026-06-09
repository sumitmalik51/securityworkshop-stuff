<#
.SYNOPSIS
    Cancel specific subscription types across ALL 271 tenants from spns.xlsx.
    Targets: Intune Suite, Defender Vuln Mgmt, M365 E5 (no Teams), Entra ID P2
#>
param(
    [string[]] $Batches   = @('Batch1','Batch2','Batch3','Batch4','Batch5'),
    [string]   $ExcelFile = 'C:\certs\spns.xlsx'
)

$ErrorActionPreference = 'Continue'
Import-Module PartnerCenter -ErrorAction Stop
Import-Module ImportExcel   -ErrorAction Stop

$targetOffers = @(
    'Microsoft Intune Suite',
    'Microsoft Defender Vulnerability Management Add-on',
    'Microsoft 365 E5 (no Teams)',
    'Microsoft Entra ID P2'
)

# Load all domains from spns.xlsx
$domains = @()
foreach ($batch in $Batches) {
    $rows = Import-Excel -Path $ExcelFile -WorksheetName $batch
    $domains += $rows | ForEach-Object { $_.orgname }
}
$domains = $domains | Select-Object -Unique
Write-Host "Loaded $($domains.Count) unique domains from spns.xlsx" -ForegroundColor Cyan

# Connect & get customers
Write-Host "Connecting to Partner Center..." -ForegroundColor Cyan
Connect-PartnerCenter

Write-Host "Getting customers..." -ForegroundColor Cyan
$allCust = Get-PartnerCustomer
$custs = $allCust | Where-Object { $domains -contains $_.Domain }
$total = $custs.Count
Write-Host "Matched $total / $($domains.Count) customers`n" -ForegroundColor Green

$i = 0
$totalCancelled = 0
$totalFailed = 0
$totalSkipped = 0

foreach ($c in $custs) {
    $i++
    $subs = Get-PartnerCustomerSubscription -CustomerId $c.CustomerId |
        Where-Object { $targetOffers -contains $_.OfferName -and $_.Status -in 'Active','Suspended' }

    if ($subs.Count -eq 0) {
        Write-Host "[$i/$total] $($c.Domain) - no target subs" -ForegroundColor DarkGray
        $totalSkipped++
        continue
    }

    Write-Host "[$i/$total] $($c.Domain) - $($subs.Count) sub(s)" -ForegroundColor White
    foreach ($s in $subs) {
        Write-Host "  $($s.Status.ToString().PadRight(10)) $($s.OfferName)" -ForegroundColor Yellow -NoNewline
        try {
            if ($s.Status -eq 'Active') {
                Set-PartnerCustomerSubscription -CustomerId $c.CustomerId `
                    -SubscriptionId $s.SubscriptionId -Status 'suspended' -ErrorAction Stop | Out-Null
            }
            Set-PartnerCustomerSubscription -CustomerId $c.CustomerId `
                -SubscriptionId $s.SubscriptionId -Status 'deleted' -ErrorAction Stop | Out-Null
            Write-Host " -> CANCELLED" -ForegroundColor Green
            $totalCancelled++
        } catch {
            Write-Host " -> FAILED: $_" -ForegroundColor Red
            $totalFailed++
        }
    }
}

Write-Host "`n=======================================" -ForegroundColor Cyan
Write-Host "DONE - $total tenants processed" -ForegroundColor Cyan
Write-Host "Cancelled: $totalCancelled" -ForegroundColor Green
Write-Host "Failed:    $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { 'Red' } else { 'Green' })
Write-Host "Skipped:   $totalSkipped (no target subs)" -ForegroundColor DarkGray
