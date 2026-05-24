<#
.SYNOPSIS
    Checks each tenant for an active "Microsoft 365 E5 (no Teams)" subscription.
    If missing/deleted, purchases a new one via Partner Center, then assigns
    the license to the odl* user via Graph API using SPN credentials.

.EXAMPLE
    # Process all tenants from SPN sheet "All" (default)
    .\Provision-M365E5.ps1

    # Process a specific SPN batch sheet
    .\Provision-M365E5.ps1 -SPNSheet "Batch1"

    # Dry run (no changes, just report)
    .\Provision-M365E5.ps1 -WhatIf

    # Process a single domain
    .\Provision-M365E5.ps1 -Domains "otuwacne104140.onmicrosoft.com"
#>

param(
    [string]$SPNFile     = "C:\certs\spns.xlsx",
    [string]$SPNSheet    = "All",
    [string]$OutputFile  = "C:\certs\Provision_M365E5_Results.csv",
    [string[]]$Domains   = @(),
    [switch]$WhatIf
)

$ErrorActionPreference = "Continue"
$offerName = "Microsoft 365 E5 (no Teams)"

# =========================================================
# INSTALL / IMPORT MODULES
# =========================================================
if (-not (Get-Module -ListAvailable -Name PartnerCenter)) {
    Write-Host "Installing PartnerCenter module..." -ForegroundColor Yellow
    Install-Module PartnerCenter -Force -Scope CurrentUser -AllowClobber
}
Import-Module PartnerCenter -ErrorAction Stop

# =========================================================
# CONNECT TO PARTNER CENTER
# =========================================================
Write-Host "Connecting to Partner Center (browser login will open)..." -ForegroundColor Cyan
Connect-PartnerCenter

# =========================================================
# LOAD TENANTS FROM SPN SHEET
# =========================================================
Write-Host "Loading tenants from $SPNFile (sheet: $SPNSheet) ..." -ForegroundColor Cyan
$allRows = Import-Excel -Path $SPNFile -WorksheetName $SPNSheet

if ($Domains.Count -gt 0) {
    $allRows = $allRows | Where-Object { $Domains -contains $_.orgname }
}

Write-Host "Total tenants to process: $($allRows.Count)`n" -ForegroundColor Cyan

# =========================================================
# NCE OFFER ID FOR M365 E5 (no Teams) - Monthly billing
# =========================================================
$nceOfferId = "CFQ7TTC0LFLZ:001L:CFQ7TTC11X36"
Write-Host "Using NCE offer: $offerName (OfferId: $nceOfferId, Monthly billing)`n" -ForegroundColor Green

# =========================================================
# GET ALL PARTNER CUSTOMERS
# =========================================================
Write-Host "Retrieving customer list..." -ForegroundColor Cyan
$allCustomers = Get-PartnerCustomer
Write-Host "Total Partner Center customers: $($allCustomers.Count)" -ForegroundColor Cyan

# Build domain -> customer lookup
$customerMap = @{}
foreach ($c in $allCustomers) {
    $customerMap[$c.Domain] = $c
}

# =========================================================
# HELPER: Get Graph token via SPN (client_credentials)
# =========================================================
function Get-GraphTokenSPN {
    param([string]$TenantId, [string]$AppId, [string]$AppSecret)
    
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $AppId
        client_secret = $AppSecret
        scope         = "https://graph.microsoft.com/.default"
    }
    try {
        $resp = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Method POST -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop
        return $resp.access_token
    } catch {
        return $null
    }
}

# =========================================================
# PROCESS EACH TENANT
# =========================================================
$results = @()
$processed = 0
$purchased = 0
$assigned  = 0
$skipped   = 0
$errors    = 0



foreach ($row in $allRows) {
    $processed++
    $domain   = $row.orgname
    $username = $row.odluser
    $tenantId = $row.TenantId
    $appId    = $row.AppId
    $appSecret = $row.AppSecret

    Write-Host "`n[$processed/$($allRows.Count)] $domain" -ForegroundColor White -NoNewline

    $result = [PSCustomObject]@{
        Domain           = $domain
        Username         = $username
        TenantId         = $tenantId
        SubStatus        = ""
        SubAction        = ""
        LicenseAssigned  = ""
        Error            = ""
    }

    # Find customer in Partner Center
    $customer = $customerMap[$domain]
    if (-not $customer) {
        Write-Host " - NOT FOUND in Partner Center" -ForegroundColor Red
        $result.Error = "Customer not found in Partner Center"
        $errors++
        $results += $result
        continue
    }

    $custId = $customer.CustomerId

    # ---- STEP 1: Check existing subscriptions ----
    try {
        $subs = Get-PartnerCustomerSubscription -CustomerId $custId -ErrorAction Stop
        $m365Sub = $subs | Where-Object { $_.OfferName -eq $offerName -and $_.Status -eq "Active" } | Select-Object -First 1
    } catch {
        Write-Host " - ERROR getting subscriptions: $($_.Exception.Message)" -ForegroundColor Red
        $result.Error = "Sub check failed: $($_.Exception.Message)"
        $errors++
        $results += $result
        continue
    }

    if ($m365Sub) {
        # Already has active M365 E5
        Write-Host " - ACTIVE (SubscriptionId: $($m365Sub.SubscriptionId))" -ForegroundColor Green
        $result.SubStatus = "Already Active"
        $result.SubAction = "None"
        $skipped++
    }
    else {
        # ---- STEP 2: Purchase new subscription ----
        $result.SubStatus = "Missing/Deleted"

        if ($WhatIf) {
            Write-Host " - WOULD PURCHASE $offerName" -ForegroundColor Yellow
            $result.SubAction = "WhatIf - would purchase"
            $skipped++
            $results += $result
            continue
        }

        Write-Host " - Purchasing (monthly)..." -ForegroundColor Yellow -NoNewline
        try {
            # Create cart with NCE offer
            $lineItem = New-Object Microsoft.Store.PartnerCenter.Models.Carts.CartLineItem
            $lineItem.CatalogItemId = $nceOfferId
            $lineItem.Quantity      = 1
            $lineItem.BillingCycle  = "Monthly"
            $lineItem.TermDuration  = "P1M"
            $cart = New-PartnerCustomerCart -CustomerId $custId -LineItems @($lineItem) -ErrorAction Stop

            # Submit the cart to complete purchase
            $order = Submit-PartnerCustomerCart -CustomerId $custId -CartId $cart.CartId -ErrorAction Stop

            Write-Host " OK (OrderId: $($order.Id))" -ForegroundColor Green
            $result.SubAction = "Purchased (Monthly)"
            $purchased++

            # Wait for provisioning
            Start-Sleep -Seconds 10
        }
        catch {
            Write-Host " FAILED: $($_.Exception.Message)" -ForegroundColor Red
            $result.SubAction = "Purchase FAILED"
            $result.Error = $_.Exception.Message
            $errors++
            $results += $result
            continue
        }
    }

    # ---- STEP 3: Assign license to user via Graph API (SPN) ----
    Write-Host "  Assigning license to $username..." -ForegroundColor Cyan -NoNewline

    if ($WhatIf) {
        Write-Host " WOULD ASSIGN" -ForegroundColor Yellow
        $result.LicenseAssigned = "WhatIf"
        $results += $result
        continue
    }

    # Get Graph token via SPN (credentials from this row)
    if (-not $appId -or -not $appSecret) {
        Write-Host " SKIPPED (no SPN creds)" -ForegroundColor Yellow
        $result.LicenseAssigned = "No SPN"
        $results += $result
        continue
    }

    $token = Get-GraphTokenSPN -TenantId $tenantId -AppId $appId -AppSecret $appSecret
    if (-not $token) {
        Write-Host " FAILED (SPN token)" -ForegroundColor Red
        $result.LicenseAssigned = "No"
        $result.Error = "Could not get Graph token via SPN"
        $errors++
        $results += $result
        continue
    }

    $headers = @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" }

    # Find the odl* user by UPN
    try {
        $userResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$username`?`$select=id,assignedLicenses" `
            -Headers $headers -ErrorAction Stop
        $userId = $userResp.id
        $existingSkus = @()
        if ($userResp.assignedLicenses) {
            $existingSkus = $userResp.assignedLicenses | Select-Object -ExpandProperty skuId
        }
    } catch {
        Write-Host " FAILED (user lookup): $($_.Exception.Message)" -ForegroundColor Red
        $result.LicenseAssigned = "No"
        $result.Error = "User lookup failed: $($_.Exception.Message)"
        $errors++
        $results += $result
        continue
    }

    # Get the actual SKU ID from tenant's subscribedSkus
    try {
        $skus = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" `
            -Headers $headers -ErrorAction Stop
        $targetSku = $skus.value | Where-Object {
            ($_.skuPartNumber -like "*SPE_E5*" -or $_.skuPartNumber -like "*M365*E5*") -and
            $_.skuPartNumber -notlike "*TEAMS*" -and
            $_.capabilityStatus -eq "Enabled"
        } | Select-Object -First 1

        if (-not $targetSku) {
            $targetSku = $skus.value | Where-Object { $_.skuPartNumber -like "*E5*" -and $_.capabilityStatus -eq "Enabled" } | Select-Object -First 1
        }
        if (-not $targetSku) {
            Write-Host " FAILED (no E5 SKU found in tenant)" -ForegroundColor Red
            $result.LicenseAssigned = "No"
            $result.Error = "No E5 SKU available in tenant yet"
            $errors++
            $results += $result
            continue
        }
        $actualSkuId = $targetSku.skuId
    } catch {
        Write-Host " FAILED (SKU lookup): $($_.Exception.Message)" -ForegroundColor Red
        $result.LicenseAssigned = "No"
        $result.Error = "SKU lookup failed: $($_.Exception.Message)"
        $errors++
        $results += $result
        continue
    }

    # Check if already assigned
    if ($existingSkus -contains $actualSkuId) {
        Write-Host " already assigned" -ForegroundColor Green
        $result.LicenseAssigned = "Already assigned"
        $results += $result
        continue
    }

    # Assign the license
    $assignBody = @{
        addLicenses    = @(@{ skuId = $actualSkuId; disabledPlans = @() })
        removeLicenses = @()
    } | ConvertTo-Json -Depth 3

    try {
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$userId/assignLicense" `
            -Method POST -Headers $headers -Body $assignBody -ErrorAction Stop | Out-Null

        Write-Host " OK" -ForegroundColor Green
        $result.LicenseAssigned = "Yes"
        $assigned++
    } catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -like "*already*" -or $errMsg -like "*conflicting*") {
            Write-Host " already assigned" -ForegroundColor Green
            $result.LicenseAssigned = "Already assigned"
        } else {
            Write-Host " FAILED: $errMsg" -ForegroundColor Red
            $result.LicenseAssigned = "No"
            $result.Error = "Assign failed: $errMsg"
            $errors++
        }
    }

    $results += $result
}

# =========================================================
# EXPORT & SUMMARY
# =========================================================
$results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Done! Processed $processed tenants." -ForegroundColor Cyan
Write-Host "Results exported to: $OutputFile" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "Already active:    $skipped" -ForegroundColor Green
Write-Host "Purchased:         $purchased" -ForegroundColor Yellow
Write-Host "Licenses assigned: $assigned" -ForegroundColor Green
Write-Host "Errors:            $errors" -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "Green" })
