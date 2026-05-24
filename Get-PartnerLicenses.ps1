<#
.SYNOPSIS
    Connects to Partner Center and retrieves license/subscription info
    for specific customer tenants. Filters by domains from your Excel file.

.NOTES
    Requires the PartnerCenter PowerShell module.
    Uses interactive browser login to authenticate.

.EXAMPLE
    # Check only tenants from Excel env 5-9
    .\Get-PartnerLicenses.ps1

    # Check tenants from a specific worksheet
    .\Get-PartnerLicenses.ps1 -Worksheet "All Environments Details"

    # Check specific domains (comma-separated)
    .\Get-PartnerLicenses.ps1 -Domains "tenant1.onmicrosoft.com","tenant2.onmicrosoft.com"

    # Check ALL partner tenants (no filter)
    .\Get-PartnerLicenses.ps1 -All
#>

param(
    [string]$InputFile   = "C:\certs\66503-User-Detail-Report-separate.xlsx",
    [string]$OutputFile  = "C:\certs\PartnerCenter_Licenses.csv",
    [string]$Worksheet   = "",
    [int[]]$Environments = @(5,6,7,8,9),
    [string[]]$Domains   = @(),
    [switch]$All
)

$ErrorActionPreference = "Continue"

# =========================================================
# INSTALL MODULE IF NEEDED
# =========================================================
if (-not (Get-Module -ListAvailable -Name PartnerCenter)) {
    Write-Host "Installing PartnerCenter module..." -ForegroundColor Yellow
    Install-Module PartnerCenter -Force -Scope CurrentUser -AllowClobber
}
Import-Module PartnerCenter -ErrorAction Stop

# =========================================================
# CONNECT TO PARTNER CENTER (interactive browser login)
# =========================================================
Write-Host "Connecting to Partner Center (browser login will open)..." -ForegroundColor Cyan
Connect-PartnerCenter

# =========================================================
# BUILD DOMAIN FILTER LIST
# =========================================================
$filterDomains = @()

if (-not $All) {
    if ($Domains.Count -gt 0) {
        # Use explicitly provided domains
        $filterDomains = $Domains
    }
    else {
        # Load domains from Excel
        if ($Worksheet -ne "") {
            $rows = Import-Excel -Path $InputFile -WorksheetName $Worksheet
            foreach ($r in $rows) {
                if ($r.username) {
                    $filterDomains += ($r.username -split "@")[1]
                }
            }
        }
        else {
            foreach ($env in $Environments) {
                $sheetName = "Environment Detail $env"
                try {
                    $rows = Import-Excel -Path $InputFile -WorksheetName $sheetName
                    foreach ($r in $rows) {
                        if ($r.username) {
                            $filterDomains += ($r.username -split "@")[1]
                        }
                    }
                } catch {
                    Write-Host "Skipping '$sheetName': $_" -ForegroundColor Yellow
                }
            }
        }
    }
    $filterDomains = $filterDomains | Select-Object -Unique
    Write-Host "Filtering to $($filterDomains.Count) tenant domains from input" -ForegroundColor Cyan
}

# =========================================================
# GET CUSTOMERS (filtered or all)
# =========================================================
Write-Host "Retrieving customer list from Partner Center..." -ForegroundColor Cyan
$allCustomers = Get-PartnerCustomer

if ($All) {
    $customers = $allCustomers
    Write-Host "Processing ALL $($customers.Count) customer tenants`n" -ForegroundColor Green
}
else {
    $customers = $allCustomers | Where-Object { $filterDomains -contains $_.Domain }
    Write-Host "Matched $($customers.Count) / $($allCustomers.Count) tenants from Partner Center`n" -ForegroundColor Green

    # Show unmatched domains
    $matchedDomains = $customers | Select-Object -ExpandProperty Domain
    $unmatched = $filterDomains | Where-Object { $matchedDomains -notcontains $_ }
    if ($unmatched.Count -gt 0) {
        Write-Host "WARNING: $($unmatched.Count) domains from input NOT found in Partner Center:" -ForegroundColor Yellow
        $unmatched | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
        Write-Host ""
    }
}

# =========================================================
# GET SUBSCRIPTIONS FOR EACH CUSTOMER
# =========================================================
$subResults = @()
$processed = 0

foreach ($customer in $customers) {
    $processed++
    $custName   = $customer.Name
    $custId     = $customer.CustomerId
    $custDomain = $customer.Domain

    Write-Host "[$processed/$($customers.Count)] $custName" -ForegroundColor White -NoNewline

    try {
        $subs = Get-PartnerCustomerSubscription -CustomerId $custId -ErrorAction Stop

        if ($subs.Count -eq 0) {
            Write-Host " - No subscriptions" -ForegroundColor Yellow
        }
        else {
            $activeSubs  = @($subs | Where-Object { $_.Status -eq "Active" })
            $deletedSubs = @($subs | Where-Object { $_.Status -eq "Deleted" })
            $otherSubs   = @($subs | Where-Object { $_.Status -notin "Active","Deleted" })

            foreach ($sub in $subs) {
                $subResults += [PSCustomObject]@{
                    CustomerName     = $custName
                    CustomerId       = $custId
                    Domain           = $custDomain
                    SubscriptionName = $sub.FriendlyName
                    SubscriptionId   = $sub.SubscriptionId
                    OfferName        = $sub.OfferName
                    Status           = $sub.Status
                    Quantity         = $sub.Quantity
                    CreationDate     = $sub.CreationDate
                    EffectiveDate    = $sub.EffectiveStartDate
                    CommitmentEnd    = $sub.CommitmentEndDate
                    AutoRenew        = $sub.AutoRenewEnabled
                }
            }
            $activeOnline = @($activeSubs | Where-Object { $_.OfferName -ne "Azure plan" })
            if ($activeOnline.Count -gt 0) {
                Write-Host " - $($activeOnline.Count) active online service(s), $($deletedSubs.Count) deleted" -ForegroundColor Green
            } elseif ($activeSubs.Count -gt 0) {
                Write-Host " - Azure plan only, $($deletedSubs.Count) online service(s) DELETED" -ForegroundColor Yellow
            } else {
                Write-Host " - ALL $($subs.Count) subscription(s) deleted/inactive" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host " - ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# =========================================================
# EXPORT RESULTS
# =========================================================
$subscriptionCsv = $OutputFile -replace '\.csv$', '_Subscriptions.csv'

$subResults | Export-Csv -Path $subscriptionCsv -NoTypeInformation -Encoding UTF8

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Done! Processed $($customers.Count) customer tenants." -ForegroundColor Cyan
Write-Host "Subscriptions exported to: $subscriptionCsv ($($subResults.Count) rows)" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

# Subscription status summary
Write-Host "`n--- Subscription Status Summary ---" -ForegroundColor Cyan
$subResults | Group-Object Status | Sort-Object Count -Descending |
    Select-Object @{N='Status';E={$_.Name}}, Count |
    Format-Table -AutoSize

# Tenants with no active online services
$tenantsNoActive = $subResults | Group-Object Domain | Where-Object {
    ($_.Group | Where-Object { $_.Status -eq "Active" -and $_.OfferName -ne "Azure plan" }).Count -eq 0
} | Select-Object -ExpandProperty Name
if ($tenantsNoActive.Count -gt 0) {
    Write-Host "Tenants with NO active online services: $($tenantsNoActive.Count)" -ForegroundColor Yellow
    $tenantsNoActive | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
}
