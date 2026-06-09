<#
.SYNOPSIS
    Purchase 3 licenses (Intune Suite, Defender Vuln Mgmt Add-on, M365 E5 no Teams)
    for each tenant that has an active odluser, then assign the licenses to that user.
    Monthly billing, 1 seat, auto-renew OFF.

.NOTES
    Requires: PartnerCenter, ImportExcel modules
    SPN in each tenant must have User.Read.All Graph permission for user lookup.
#>
param(
    [string[]] $Batches   = @('Batch1','Batch2','Batch3','Batch4','Batch5'),
    [string]   $ExcelFile = 'C:\certs\spns.xlsx'
)

$ErrorActionPreference = 'Continue'
Import-Module PartnerCenter -ErrorAction Stop
Import-Module ImportExcel   -ErrorAction Stop

# ── Catalog Item IDs (NCE monthly P1M) ──────────────────
$catalogItems = @(
   # @{ Name = 'Microsoft Intune Suite';                              CatalogItemId = 'CFQ7TTC0RZFJ:0001:CFQ7TTC0S5K1' },
    #@{ Name = 'Microsoft Defender Vulnerability Management Add-on';  CatalogItemId = 'CFQ7TTC0JPGV:0002:CFQ7TTC08V1M' },
    @{ Name = 'Microsoft 365 E5 (no Teams)';                         CatalogItemId = 'CFQ7TTC0LFLZ:001L:CFQ7TTC11X36' }
)

# ── License SKU IDs for assignment ───────────────────────
$licenseSkuIds = @(
    #'a929cd4d-8672-47c9-8664-159c1f322ba8',   # Microsoft Intune Suite
    #'ad7a56e0-6903-4d13-94f3-5ad491e78960',   # Defender Vuln Mgmt Add-on
    '18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e'    # M365 E5 (no Teams)
)

# ── Load accounts from spns.xlsx ─────────────────────────
$accounts = @()
foreach ($batch in $Batches) {
    $accounts += Import-Excel -Path $ExcelFile -WorksheetName $batch
}
$total = $accounts.Count
Write-Host "Loaded $total accounts from spns.xlsx" -ForegroundColor Cyan

# ── Connect Partner Center ───────────────────────────────
Write-Host "Connecting to Partner Center..." -ForegroundColor Cyan
Connect-PartnerCenter
$allCust = Get-PartnerCustomer

# ── Helper: check if user exists via Graph ───────────────
function Test-UserExists {
    param([string]$UPN, [string]$TenantId, [string]$AppId, [string]$AppSecret)
    try {
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $AppId
            client_secret = $AppSecret
            scope         = 'https://graph.microsoft.com/.default'
        }
        $tok = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -ContentType 'application/x-www-form-urlencoded' -Body $body -ErrorAction Stop
        $headers = @{ Authorization = "Bearer $($tok.access_token)" }
        $filter = "userPrincipalName eq '$UPN'"
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=id,userPrincipalName&`$top=1"
        $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
        if ($resp.value.Count -gt 0) { return $resp.value[0].id }
        return $null
    } catch {
        Write-Host "    Graph error: $_" -ForegroundColor DarkYellow
        return $null
    }
}

# ── Process each tenant ──────────────────────────────────
$stats = @{ Purchased = 0; Assigned = 0; Skipped = 0; NoUser = 0; Failed = 0; AlreadyHas = 0 }

for ($i = 0; $i -lt $total; $i++) {
    $acct = $accounts[$i]
    $idx  = $i + 1
    $domain = $acct.orgname
    $short  = ($domain -split '\.')[0]

    Write-Host "`n[$idx/$total] $short" -ForegroundColor Cyan -NoNewline

    # Find Partner Center customer
    $cust = $allCust | Where-Object { $_.Domain -eq $domain }
    if (-not $cust) {
        Write-Host " - NOT FOUND in Partner Center" -ForegroundColor Red
        $stats.Failed++
        continue
    }

    # Check if odluser exists
    $userId = Test-UserExists -UPN $acct.odluser -TenantId $acct.TenantId `
        -AppId $acct.AppId -AppSecret $acct.AppSecret
    if (-not $userId) {
        Write-Host " - no odluser, skipping" -ForegroundColor DarkGray
        $stats.NoUser++
        continue
    }
    Write-Host "" # newline after tenant name

    # Check existing subscriptions to avoid re-purchasing
    $existingSubs = Get-PartnerCustomerSubscription -CustomerId $cust.CustomerId |
        Where-Object { $_.Status -eq 'Active' }
    $existingOffers = $existingSubs | ForEach-Object { $_.OfferName }

    # Check user's existing license assignments
    $userLicenses = Get-PartnerCustomerUserLicense -CustomerId $cust.CustomerId -UserId $userId -ErrorAction SilentlyContinue
    $userSkuIds = @()
    if ($userLicenses) { $userSkuIds = $userLicenses | ForEach-Object { $_.SkuId } }

    # If user already has all 3 licenses, skip entirely
    $allAssigned = $true
    foreach ($skuId in $licenseSkuIds) {
        if ($userSkuIds -notcontains $skuId) { $allAssigned = $false; break }
    }
    if ($allAssigned -and ($existingOffers | Where-Object { ($catalogItems | ForEach-Object { $_.Name }) -contains $_ }).Count -ge $catalogItems.Count) {
        Write-Host " - already has all subs + licenses, skipping" -ForegroundColor DarkGray
        $stats.AlreadyHas += $catalogItems.Count
        continue
    }

    Write-Host "" # newline after tenant name

    $itemsToBuy = @()
    $lineItemId = 0
    foreach ($item in $catalogItems) {
        if ($existingOffers -contains $item.Name) {
            Write-Host "  Already has sub: $($item.Name)" -ForegroundColor DarkGray
            $stats.AlreadyHas++
        } else {
            $lineItem = New-Object -TypeName Microsoft.Store.PartnerCenter.PowerShell.Models.Carts.PSCartLineItem
            $lineItem.Id = $lineItemId++
            $lineItem.BillingCycle = 'Monthly'
            $lineItem.CatalogItemId = $item.CatalogItemId
            $lineItem.Quantity = 1
            $lineItem.TermDuration = 'P1M'
            $itemsToBuy += $lineItem
        }
    }

    # Purchase via cart
    if ($itemsToBuy.Count -gt 0) {
        try {
            $cart = New-PartnerCustomerCart -CustomerId $cust.CustomerId -LineItems $itemsToBuy -ErrorAction Stop
            # Check for line item errors before submitting
            $hasErrors = $false
            foreach ($li in $cart.LineItems) {
                if ($li.Error) {
                    Write-Host "  Cart error on $($li.CatalogItemId): $($li.Error.Message)" -ForegroundColor Red
                    $hasErrors = $true
                }
            }
            if ($hasErrors) {
                Write-Host "  Skipping purchase due to cart errors" -ForegroundColor Red
                $stats.Failed++
                continue
            }
            $order = Submit-PartnerCustomerCart -CustomerId $cust.CustomerId -CartId $cart.CartId -ErrorAction Stop
            Write-Host "  Purchased $($itemsToBuy.Count) subscription(s)" -ForegroundColor Green
            $stats.Purchased += $itemsToBuy.Count
        } catch {
            Write-Host "  Purchase FAILED: $_" -ForegroundColor Red
            $stats.Failed++
            continue
        }

        # Brief pause for provisioning
        Start-Sleep -Seconds 5
    }

    # Assign only licenses the user doesn't already have
    try {
        $skus = Get-PartnerCustomerSubscribedSku -CustomerId $cust.CustomerId
        $licUpdate = New-Object -TypeName Microsoft.Store.PartnerCenter.PowerShell.Models.Licenses.PSLicenseUpdate

        foreach ($skuId in $licenseSkuIds) {
            if ($userSkuIds -contains $skuId) {
                Write-Host "  Already assigned: $($skuId)" -ForegroundColor DarkGray
                continue
            }
            $matchedSku = $skus | Where-Object { $_.SkuId -eq $skuId }
            if ($matchedSku) {
                $lic = New-Object -TypeName Microsoft.Store.PartnerCenter.PowerShell.Models.Licenses.PSLicenseAssignment
                $lic.SkuId = $skuId
                $licUpdate.LicensesToAssign.Add($lic)
            }
        }

        if ($licUpdate.LicensesToAssign.Count -gt 0) {
            Set-PartnerCustomerUserLicense -CustomerId $cust.CustomerId `
                -UserId $userId -LicenseUpdate $licUpdate -ErrorAction Stop | Out-Null
            Write-Host "  Assigned $($licUpdate.LicensesToAssign.Count) license(s) to $($acct.odluser)" -ForegroundColor Green
            $stats.Assigned += $licUpdate.LicensesToAssign.Count
        } else {
            Write-Host "  All licenses already assigned" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "  License assign FAILED: $_" -ForegroundColor Red
        $stats.Failed++
    }
}

Write-Host "`n=======================================" -ForegroundColor Cyan
Write-Host "DONE - $total tenants processed" -ForegroundColor Cyan
Write-Host "Purchased:  $($stats.Purchased)" -ForegroundColor Green
Write-Host "Assigned:   $($stats.Assigned)" -ForegroundColor Green
Write-Host "Already had: $($stats.AlreadyHas)" -ForegroundColor DarkGray
Write-Host "No user:    $($stats.NoUser)" -ForegroundColor DarkGray
Write-Host "Skipped:    $($stats.Skipped)" -ForegroundColor DarkGray
Write-Host "Failed:     $($stats.Failed)" -ForegroundColor $(if ($stats.Failed -gt 0) { 'Red' } else { 'Green' })
