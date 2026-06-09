<#
.SYNOPSIS
    Check and cancel specific subscription types for the 62 orphaned tenants.
    Targets: Intune Suite, Defender Vuln Mgmt, M365 E5 (no Teams), Entra ID P2
#>
param()

$ErrorActionPreference = 'Continue'
Import-Module PartnerCenter -ErrorAction Stop

$domains = @(
    'otuwacne104045.onmicrosoft.com','otuwacne104046.onmicrosoft.com','otuwacne104047.onmicrosoft.com',
    'otuwacne104054.onmicrosoft.com','otuwacne104055.onmicrosoft.com','otuwacne104056.onmicrosoft.com',
    'otuwacne104057.onmicrosoft.com','otuwacne104058.onmicrosoft.com','otuwacne104060.onmicrosoft.com',
    'otuwacne104061.onmicrosoft.com','otuwacne104062.onmicrosoft.com','otuwacne104063.onmicrosoft.com',
    'otuwacne104064.onmicrosoft.com','otuwacne104041.onmicrosoft.com','otuwacne104042.onmicrosoft.com',
    'otuwacne104043.onmicrosoft.com','otuwacne104044.onmicrosoft.com','otuwacne104052.onmicrosoft.com',
    'otuwacne104051.onmicrosoft.com','otuwacne104053.onmicrosoft.com','otuwacne104048.onmicrosoft.com',
    'otuwacne104065.onmicrosoft.com','otuwacne104050.onmicrosoft.com','otuwacne103923.onmicrosoft.com',
    'otuwacne103926.onmicrosoft.com','otuwacne103982.onmicrosoft.com','otuwacne103991.onmicrosoft.com',
    'otuwacne104127.onmicrosoft.com','otuwacne104092.onmicrosoft.com','otuwacne104080.onmicrosoft.com',
    'otuwacne104099.onmicrosoft.com','otuwacne104131.onmicrosoft.com','otuwacne104132.onmicrosoft.com',
    'otuwacne104133.onmicrosoft.com','otuwacne104134.onmicrosoft.com','otuwacne104135.onmicrosoft.com',
    'otuwacne104136.onmicrosoft.com','otuwacne104137.onmicrosoft.com','otuwacne104138.onmicrosoft.com',
    'otuwacne104139.onmicrosoft.com','otuwacne104140.onmicrosoft.com','otuwacne103892.onmicrosoft.com',
    'otuwacne103897.onmicrosoft.com','otuwacne103900.onmicrosoft.com','otuwacne103902.onmicrosoft.com',
    'otuwacne103906.onmicrosoft.com','otuwacne103930.onmicrosoft.com','otuwacne103937.onmicrosoft.com',
    'otuwacne103946.onmicrosoft.com','otuwacne103947.onmicrosoft.com','otuwacne103950.onmicrosoft.com',
    'otuwacne103956.onmicrosoft.com','otuwacne103957.onmicrosoft.com','otuwacne103958.onmicrosoft.com',
    'otuwacne103960.onmicrosoft.com','otuwacne103971.onmicrosoft.com','otuwacne103974.onmicrosoft.com',
    'otuwacne103977.onmicrosoft.com','otuwacne103989.onmicrosoft.com','otuwacne103999.onmicrosoft.com',
    'otuwacne104009.onmicrosoft.com','otuwacne104119.onmicrosoft.com'
)

$targetOffers = @(
    'Microsoft Intune Suite',
    'Microsoft Defender Vulnerability Management Add-on',
    'Microsoft 365 E5 (no Teams)',
    'Microsoft Entra ID P2'
)

Write-Host "Connecting to Partner Center..." -ForegroundColor Cyan
Connect-PartnerCenter

Write-Host "Getting customers..." -ForegroundColor Cyan
$allCust = Get-PartnerCustomer
$custs = $allCust | Where-Object { $domains -contains $_.Domain }
Write-Host "Matched $($custs.Count) / $($domains.Count) customers`n" -ForegroundColor Green

$i = 0
$totalCancelled = 0
$totalFailed = 0
$totalSkipped = 0

foreach ($c in $custs) {
    $i++
    $subs = Get-PartnerCustomerSubscription -CustomerId $c.CustomerId |
        Where-Object { $targetOffers -contains $_.OfferName -and $_.Status -in 'Active','Suspended' }

    if ($subs.Count -eq 0) {
        Write-Host "[$i/62] $($c.Domain) - no target subs" -ForegroundColor DarkGray
        $totalSkipped++
        continue
    }

    Write-Host "[$i/62] $($c.Domain) - $($subs.Count) sub(s)" -ForegroundColor White
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
Write-Host "DONE" -ForegroundColor Cyan
Write-Host "Cancelled: $totalCancelled" -ForegroundColor Green
Write-Host "Failed:    $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { 'Red' } else { 'Green' })
Write-Host "Skipped:   $totalSkipped (no target subs)" -ForegroundColor DarkGray
