<#
.SYNOPSIS
    Assigns M365 E5 (no Teams) license to odl users and verifies assignment.
    Uses SPN credentials from spns.xlsx (same as Provision-M365E5.ps1).
    No Partner Center needed — purely Graph API via SPN client_credentials.

.EXAMPLE
    .\Assign-VerifyLicense.ps1
    .\Assign-VerifyLicense.ps1 -SPNSheet "Batch1" -OutputFile "C:\certs\LicenseVerify_Batch1.csv"
#>

param(
    [string]$SPNFile    = "C:\certs\spns.xlsx",
    [string]$SPNSheet   = "All",
    [string]$OutputFile = "C:\certs\LicenseVerify_Results.csv",
    [string[]]$Domains  = @()
)

$ErrorActionPreference = "Continue"

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
$results   = @()
$processed = 0
$assigned  = 0
$already   = 0
$failed    = 0

foreach ($row in $allRows) {
    $processed++
    $domain    = $row.orgname
    $username  = $row.odluser
    $tenantId  = $row.TenantId
    $appId     = $row.AppId
    $appSecret = $row.AppSecret

    Write-Host "[$processed/$($allRows.Count)] $domain" -ForegroundColor White -NoNewline

    $result = [PSCustomObject]@{
        Domain          = $domain
        Username        = $username
        TenantId        = $tenantId
        SkuPartNumber   = ""
        SkuId           = ""
        LicenseStatus   = ""
        Error           = ""
    }

    # -- Get Graph token --
    if (-not $appId -or -not $appSecret) {
        Write-Host " - SKIP (no SPN creds)" -ForegroundColor Yellow
        $result.LicenseStatus = "Skipped"
        $result.Error = "No SPN credentials"
        $failed++
        $results += $result
        continue
    }

    $token = Get-GraphTokenSPN -TenantId $tenantId -AppId $appId -AppSecret $appSecret
    if (-not $token) {
        Write-Host " - FAIL (token)" -ForegroundColor Red
        $result.LicenseStatus = "TokenFailed"
        $result.Error = "Could not get Graph token"
        $failed++
        $results += $result
        continue
    }

    $headers = @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" }

    # -- Find E5 SKU in tenant --
    try {
        $skus = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" `
            -Headers $headers -ErrorAction Stop
        $targetSku = $skus.value | Where-Object {
            ($_.skuPartNumber -like "*SPE_E5*" -or $_.skuPartNumber -like "*M365*E5*") -and
            $_.skuPartNumber -notlike "*TEAMS*" -and
            $_.capabilityStatus -eq "Enabled"
        } | Select-Object -First 1

        if (-not $targetSku) {
            $targetSku = $skus.value | Where-Object {
                $_.skuPartNumber -like "*E5*" -and $_.capabilityStatus -eq "Enabled"
            } | Select-Object -First 1
        }

        if (-not $targetSku) {
            Write-Host " - NO E5 SKU in tenant" -ForegroundColor Red
            $result.LicenseStatus = "NoSKU"
            $result.Error = "No E5 SKU found (not yet provisioned?)"
            $failed++
            $results += $result
            continue
        }

        $result.SkuPartNumber = $targetSku.skuPartNumber
        $result.SkuId = $targetSku.skuId
    } catch {
        Write-Host " - FAIL (SKU lookup): $($_.Exception.Message)" -ForegroundColor Red
        $result.LicenseStatus = "SKUFailed"
        $result.Error = $_.Exception.Message
        $failed++
        $results += $result
        continue
    }

    # -- Lookup user and check current licenses --
    try {
        $userResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$username`?`$select=id,assignedLicenses,displayName" `
            -Headers $headers -ErrorAction Stop
        $userId = $userResp.id
        $existingSkus = @()
        if ($userResp.assignedLicenses) {
            $existingSkus = $userResp.assignedLicenses | Select-Object -ExpandProperty skuId
        }
    } catch {
        Write-Host " - FAIL (user lookup): $($_.Exception.Message)" -ForegroundColor Red
        $result.LicenseStatus = "UserNotFound"
        $result.Error = $_.Exception.Message
        $failed++
        $results += $result
        continue
    }

    # -- Check if already assigned --
    if ($existingSkus -contains $targetSku.skuId) {
        Write-Host " - already assigned" -ForegroundColor Green
        $result.LicenseStatus = "Assigned"
        $already++
        $results += $result
        continue
    }

    # -- Assign the license --
    $assignBody = @{
        addLicenses    = @(@{ skuId = $targetSku.skuId; disabledPlans = @() })
        removeLicenses = @()
    } | ConvertTo-Json -Depth 3

    try {
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$userId/assignLicense" `
            -Method POST -Headers $headers -Body $assignBody -ErrorAction Stop | Out-Null

        Write-Host " - ASSIGNED" -ForegroundColor Green
        $result.LicenseStatus = "Assigned"
        $assigned++
    } catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -like "*already*" -or $errMsg -like "*conflicting*") {
            Write-Host " - already assigned" -ForegroundColor Green
            $result.LicenseStatus = "Assigned"
            $already++
        } else {
            Write-Host " - FAIL (assign): $errMsg" -ForegroundColor Red
            $result.LicenseStatus = "AssignFailed"
            $result.Error = $errMsg
            $failed++
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

$assignedCount = ($results | Where-Object { $_.LicenseStatus -eq "Assigned" }).Count
$noSkuCount    = ($results | Where-Object { $_.LicenseStatus -eq "NoSKU" }).Count
$failCount     = ($results | Where-Object { $_.LicenseStatus -notin @("Assigned","Skipped") -and $_.LicenseStatus -ne "" }).Count

Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "Total:             $processed" -ForegroundColor White
Write-Host "Licensed (OK):     $assignedCount" -ForegroundColor Green
Write-Host "Newly assigned:    $assigned" -ForegroundColor Yellow
Write-Host "Already assigned:  $already" -ForegroundColor Green
Write-Host "No E5 SKU yet:     $noSkuCount" -ForegroundColor $(if ($noSkuCount -gt 0) { "Red" } else { "Green" })
Write-Host "Other failures:    $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
