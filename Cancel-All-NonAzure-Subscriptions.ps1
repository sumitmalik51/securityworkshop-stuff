<#
.SYNOPSIS
    Cancels all non-Azure subscriptions for all Partner Center customers.

.DESCRIPTION
    For every customer in Partner Center:
      1. Gets subscriptions in Active or Suspended status
      2. Excludes Azure subscriptions/offers
      3. Cancels each target subscription (Active -> Suspended -> Deleted)

.EXAMPLE
    .\Cancel-All-NonAzure-Subscriptions.ps1
    .\Cancel-All-NonAzure-Subscriptions.ps1 -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]]$Batches = @('Batch1','Batch2','Batch3','Batch4','Batch5'),
    [string]$ExcelFile = 'C:\certs\spns.xlsx',
    [switch]$UseDeviceAuthentication,
    [string]$OutputCsv = 'C:\certs\Cancel-All-NonAzure-Results.csv'
)

$ErrorActionPreference = 'Continue'

if (-not (Get-Module -ListAvailable -Name PartnerCenter)) {
    Write-Host "Installing PartnerCenter module..." -ForegroundColor Yellow
    Install-Module PartnerCenter -Force -Scope CurrentUser -AllowClobber
}
Import-Module PartnerCenter -ErrorAction Stop

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module ImportExcel -Force -Scope CurrentUser -AllowClobber
}
Import-Module ImportExcel -ErrorAction Stop

function Test-IsAzureOffer {
    param([string]$OfferName)
    if ([string]::IsNullOrWhiteSpace($OfferName)) { return $false }

    $name = $OfferName.ToLowerInvariant().Trim()

    # Keep Azure offers untouched.
    if ($name -eq 'azure plan') { return $true }
    if ($name -eq 'microsoft azure') { return $true }
    if ($name -like '*azure reservation*') { return $true }
    if ($name -like '*savings plan*azure*') { return $true }

    return $false
}

Write-Host "Connecting to Partner Center..." -ForegroundColor Cyan
try {
    if ($UseDeviceAuthentication) {
        Connect-PartnerCenter -UseDeviceAuthentication -ErrorAction Stop
    } else {
        Connect-PartnerCenter -ErrorAction Stop
    }
} catch {
    Write-Host "Partner Center login failed: $_" -ForegroundColor Red
    throw
}

# Load target domains from SPN sheet
$domains = @()
foreach ($batch in $Batches) {
    $rows = Import-Excel -Path $ExcelFile -WorksheetName $batch -ErrorAction Stop
    $domains += $rows | ForEach-Object { $_.orgname }
}
$domains = @($domains | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
Write-Host "Loaded $($domains.Count) unique domains from spns.xlsx" -ForegroundColor Green

Write-Host "Getting all customers..." -ForegroundColor Cyan
$allCustomers = @(Get-PartnerCustomer)
$customers = @($allCustomers | Where-Object { $domains -contains $_.Domain })
$totalCustomers = $customers.Count
Write-Host "Matched $totalCustomers / $($domains.Count) customers" -ForegroundColor Green

$matchedDomains = @($customers | Select-Object -ExpandProperty Domain)
$unmatched = @($domains | Where-Object { $matchedDomains -notcontains $_ })
if ($unmatched.Count -gt 0) {
    Write-Host "Domains not found in Partner Center: $($unmatched.Count)" -ForegroundColor Yellow
}

$results = @()
$totalCancelled = 0
$totalFailed = 0
$totalSkippedNoTargets = 0
$totalAzureSkipped = 0

for ($i = 0; $i -lt $totalCustomers; $i++) {
    $cust = $customers[$i]
    $idx = $i + 1

    Write-Host "`n=== [$idx/$totalCustomers] $($cust.Domain) ===" -ForegroundColor Cyan

    try {
        $subs = @(Get-PartnerCustomerSubscription -CustomerId $cust.CustomerId -ErrorAction Stop)
        $activeOrSuspended = @($subs | Where-Object { $_.Status -in 'Active', 'Suspended' })

        if ($activeOrSuspended.Count -eq 0) {
            Write-Host "  No Active/Suspended subscriptions" -ForegroundColor DarkGray
            $totalSkippedNoTargets++
            $results += [pscustomobject]@{
                Domain = $cust.Domain
                CustomerId = $cust.CustomerId
                Action = 'SKIP'
                Cancelled = 0
                Failed = 0
                AzureSkipped = 0
                Details = 'No Active/Suspended subscriptions'
            }
            continue
        }

        $azureSubs = @($activeOrSuspended | Where-Object { Test-IsAzureOffer $_.OfferName })
        $targets = @($activeOrSuspended | Where-Object { -not (Test-IsAzureOffer $_.OfferName) })
        $totalAzureSkipped += $azureSubs.Count

        if ($targets.Count -eq 0) {
            Write-Host "  Only Azure subscriptions found; nothing to cancel" -ForegroundColor DarkGray
            $totalSkippedNoTargets++
            $results += [pscustomobject]@{
                Domain = $cust.Domain
                CustomerId = $cust.CustomerId
                Action = 'SKIP'
                Cancelled = 0
                Failed = 0
                AzureSkipped = $azureSubs.Count
                Details = 'Only Azure subscriptions'
            }
            continue
        }

        Write-Host "  Cancelling $($targets.Count) non-Azure subscription(s); skipping $($azureSubs.Count) Azure subscription(s)" -ForegroundColor White

        $custCancelled = 0
        $custFailed = 0

        foreach ($sub in $targets) {
            $subLabel = "$($sub.FriendlyName) ($($sub.OfferName))"
            $subStatus = [string]$sub.Status

            Write-Host "    $($subStatus.PadRight(10)) $subLabel" -ForegroundColor Yellow -NoNewline

            if ($PSCmdlet.ShouldProcess("$($cust.Domain) :: $subLabel", 'Cancel subscription')) {
                try {
                    if ($subStatus -eq 'Active') {
                        Set-PartnerCustomerSubscription -CustomerId $cust.CustomerId `
                            -SubscriptionId $sub.SubscriptionId -Status 'suspended' -ErrorAction Stop | Out-Null
                    }

                    Set-PartnerCustomerSubscription -CustomerId $cust.CustomerId `
                        -SubscriptionId $sub.SubscriptionId -Status 'deleted' -ErrorAction Stop | Out-Null

                    Write-Host " -> CANCELLED" -ForegroundColor Green
                    $custCancelled++
                    $totalCancelled++
                }
                catch {
                    Write-Host " -> FAILED: $_" -ForegroundColor Red
                    $custFailed++
                    $totalFailed++
                }
            }
        }

        $results += [pscustomobject]@{
            Domain = $cust.Domain
            CustomerId = $cust.CustomerId
            Action = if ($custFailed -gt 0) { 'PARTIAL/ERROR' } else { 'CANCELLED' }
            Cancelled = $custCancelled
            Failed = $custFailed
            AzureSkipped = $azureSubs.Count
            Details = "Targets=$($targets.Count); AzureSkipped=$($azureSubs.Count)"
        }

        Write-Host "  Done - cancelled: $custCancelled, failed: $custFailed" -ForegroundColor Magenta
    }
    catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
        $totalFailed++
        $results += [pscustomobject]@{
            Domain = $cust.Domain
            CustomerId = $cust.CustomerId
            Action = 'ERROR'
            Cancelled = 0
            Failed = 1
            AzureSkipped = 0
            Details = "$_"
        }
    }
}

Write-Host "`n=======================================" -ForegroundColor Cyan
Write-Host "DONE" -ForegroundColor Cyan
Write-Host "Customers processed: $totalCustomers" -ForegroundColor Cyan
Write-Host "Cancelled:          $totalCancelled" -ForegroundColor Green
Write-Host "Failed:             $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { 'Red' } else { 'Green' })
Write-Host "Azure skipped:      $totalAzureSkipped" -ForegroundColor DarkGray
Write-Host "No-target customers:$totalSkippedNoTargets" -ForegroundColor DarkGray

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Force
Write-Host "Results saved to: $OutputCsv" -ForegroundColor Cyan
