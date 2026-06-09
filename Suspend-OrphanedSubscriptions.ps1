<#
.SYNOPSIS
    Suspends all active subscriptions for orphaned tenants via Partner Center.

.DESCRIPTION
    Uses interactive browser login to Partner Center, then for each target domain:
      1. Looks up the customer by domain
      2. Gets all active subscriptions
      3. Suspends each subscription (sets Status = "suspended")

.EXAMPLE
    .\Suspend-OrphanedSubscriptions.ps1
    .\Suspend-OrphanedSubscriptions.ps1 -WhatIf
    .\Suspend-OrphanedSubscriptions.ps1 -Domains "otuwacne103923.onmicrosoft.com"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]]$Domains = @(
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
    ),
    [string]$OutputCsv = 'C:\certs\Suspended-Subscriptions.csv'
)

$ErrorActionPreference = 'Continue'

# ── ensure module ──
if (-not (Get-Module -ListAvailable -Name PartnerCenter)) {
    Write-Host "Installing PartnerCenter module..." -ForegroundColor Yellow
    Install-Module PartnerCenter -Force -Scope CurrentUser -AllowClobber
}
Import-Module PartnerCenter -ErrorAction Stop

# ── connect (browser login) ──
Write-Host "Connecting to Partner Center (browser login)..." -ForegroundColor Cyan
Connect-PartnerCenter

# ── get all customers ──
Write-Host "Retrieving customer list..." -ForegroundColor Cyan
$allCustomers = Get-PartnerCustomer
$customers = $allCustomers | Where-Object { $Domains -contains $_.Domain }
Write-Host "Matched $($customers.Count) / $($Domains.Count) target domains`n" -ForegroundColor Green

# show unmatched
$matchedDomains = $customers | Select-Object -ExpandProperty Domain
$unmatched = $Domains | Where-Object { $matchedDomains -notcontains $_ }
if ($unmatched.Count -gt 0) {
    Write-Host "Not found in Partner Center ($($unmatched.Count)):" -ForegroundColor Yellow
    $unmatched | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkYellow }
    Write-Host ""
}

# ── process each customer ──
$results = @()
$processed = 0
$totalSuspended = 0

foreach ($customer in $customers) {
    $processed++
    $custName   = $customer.Name
    $custId     = $customer.CustomerId
    $custDomain = $customer.Domain

    Write-Host "=== [$processed/$($customers.Count)] $custDomain ===" -ForegroundColor Cyan

    try {
        $subs = Get-PartnerCustomerSubscription -CustomerId $custId -ErrorAction Stop
        $targetSubs = @($subs | Where-Object { $_.Status -in 'Active','Suspended' -and $_.OfferName -ne 'Azure plan' })

        if ($targetSubs.Count -eq 0) {
            Write-Host "  No active/suspended subscriptions to cancel" -ForegroundColor DarkGray
            $results += [pscustomobject]@{
                Domain=$custDomain; Customer=$custName; CustomerId=$custId
                Action="SKIP"; ActiveSubs=0; Suspended=0; Details="Nothing to cancel"
            }
            continue
        }

        Write-Host "  Found $($targetSubs.Count) subscription(s) to cancel" -ForegroundColor White
        $cancelCount = 0

        foreach ($sub in $targetSubs) {
            $subName = $sub.FriendlyName
            $subId   = $sub.SubscriptionId
            $curStatus = $sub.Status
            Write-Host "    Cancelling: $subName [$curStatus] ($($sub.OfferName))" -ForegroundColor Yellow

            if ($PSCmdlet.ShouldProcess("$custDomain - $subName", "Cancel subscription")) {
                try {
                    # Suspend first if active, then delete
                    if ($curStatus -eq 'Active') {
                        Set-PartnerCustomerSubscription -CustomerId $custId `
                            -SubscriptionId $subId -Status 'suspended' -ErrorAction Stop | Out-Null
                    }
                    Set-PartnerCustomerSubscription -CustomerId $custId `
                        -SubscriptionId $subId -Status 'deleted' -ErrorAction Stop | Out-Null
                    Write-Host "      CANCELLED" -ForegroundColor Green
                    $cancelCount++
                } catch {
                    Write-Host "      FAILED: $_" -ForegroundColor Red
                }
            }
        }

        $totalSuspended += $cancelCount
        $results += [pscustomobject]@{
            Domain=$custDomain; Customer=$custName; CustomerId=$custId
            Action="CANCELLED"; ActiveSubs=$targetSubs.Count; Suspended=$cancelCount
            Details=($targetSubs | ForEach-Object { "$($_.FriendlyName) [$($_.Status)]" }) -join '; '
        }
        Write-Host "  Done - $cancelCount/$($targetSubs.Count) cancelled" -ForegroundColor Magenta

    } catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
        $results += [pscustomobject]@{
            Domain=$custDomain; Customer=$custName; CustomerId=$custId
            Action="ERROR"; ActiveSubs=0; Suspended=0; Details="$_"
        }
    }
}

# ── summary ──
Write-Host "`n=======================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "Customers processed: $($customers.Count)"
Write-Host "Total subscriptions cancelled: $totalSuspended"
Write-Host ""
$results | Format-Table -AutoSize

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Force
Write-Host "Results saved to $OutputCsv"
