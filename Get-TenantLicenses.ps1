<#
.SYNOPSIS
    Checks assigned licenses across all tenant environments and exports to CSV.

.NOTES
    Reads credentials from the Excel file and connects via MSOnline to retrieve
    subscribed SKUs (licenses) per tenant. Exports results to CSV.
#>

param(
    [string]$InputFile  = "C:\certs\66503-User-Detail-Report-separate.xlsx",
    [string]$OutputFile = "C:\certs\TenantLicenses.csv",
    [int[]]$Environments = @(1,2,3,4,5,6,7,8,9),
    [string]$Worksheet  = ""
)

$ErrorActionPreference = "Continue"

# =========================================================
# COLLECT ALL ACCOUNTS FROM SPECIFIED ENVIRONMENTS
# =========================================================
$allRows = @()

if ($Worksheet -ne "") {
    # Load from a single named worksheet
    try {
        $rows = Import-Excel -Path $InputFile -WorksheetName $Worksheet
        foreach ($r in $rows) {
            $r | Add-Member -NotePropertyName "Environment" -NotePropertyValue $Worksheet -Force
        }
        $allRows += $rows
        Write-Host "Loaded $($rows.Count) tenants from '$Worksheet'" -ForegroundColor Cyan
    } catch {
        Write-Host "Failed to load '$Worksheet': $_" -ForegroundColor Red
        return
    }
} else {
    foreach ($env in $Environments) {
        $sheetName = "Environment Detail $env"
        try {
            $rows = Import-Excel -Path $InputFile -WorksheetName $sheetName
            foreach ($r in $rows) {
                $r | Add-Member -NotePropertyName "Environment" -NotePropertyValue $env -Force
            }
            $allRows += $rows
            Write-Host "Loaded $($rows.Count) tenants from '$sheetName'" -ForegroundColor Cyan
        } catch {
            Write-Host "Skipping '$sheetName': $_" -ForegroundColor Yellow
        }
    }
}

Write-Host "`nTotal tenants to process: $($allRows.Count)`n"

# =========================================================
# PROCESS EACH TENANT
# =========================================================
$results = @()
$processed = 0

foreach ($row in $allRows) {

    $user     = $row.username
    $pass     = $row.password
    $tenantId = $row.'tenant id'

    if (-not $user -or -not $pass) {
        Write-Host "Skipping row with missing credentials" -ForegroundColor Yellow
        continue
    }

    $processed++
    $domain = $user.Split("@")[1]
    Write-Host "[$processed/$($allRows.Count)] Processing: $user" -ForegroundColor White

    # Get token via ROPC flow and call Graph API directly
    try {
        $tokenBody = @{
            grant_type = "password"
            client_id  = "1b730954-1685-4b74-9bfd-dac224a7b894"  # Azure AD PowerShell well-known client
            scope      = "https://graph.microsoft.com/.default"
            username   = $user
            password   = $pass
        }
        $tenantDomain = $user.Split("@")[1]
        $tokenUrl = "https://login.microsoftonline.com/$tenantDomain/oauth2/v2.0/token"
        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $tokenBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
        $accessToken = $tokenResponse.access_token

        $headers = @{ Authorization = "Bearer $accessToken" }
        $skuResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $headers -ErrorAction Stop
        $skus = $skuResponse.value
    } catch {
        $skus = $null
        $licenseError = $_.Exception.Message
    }

    $licenseOutput = @()
    if ($licenseError) {
        $licenseOutput += "ERROR|$licenseError"
    } elseif (-not $skus -or $skus.Count -eq 0) {
        $licenseOutput += "NO_LICENSES_FOUND"
    } else {
        foreach ($sku in $skus) {
            $skuName   = $sku.skuPartNumber
            $active    = $sku.prepaidUnits.enabled
            $consumed  = $sku.consumedUnits
            $warning   = $sku.prepaidUnits.warning
            $suspended = $sku.prepaidUnits.suspended
            $licenseOutput += "SKU|$skuName|$active|$consumed|$warning|$suspended"
        }
    }
    $licenseError = $null

    $foundLicenses = $false

    foreach ($line in $licenseOutput) {
        if ($line -match "^SKU\|(.+)$") {
            $parts = $Matches[1] -split "\|"
            $results += [PSCustomObject]@{
                Environment   = $row.Environment
                Username      = $user
                TenantId      = $tenantId
                LicenseSKU    = $parts[0]
                ActiveUnits   = $parts[1]
                ConsumedUnits = $parts[2]
                WarningUnits  = $parts[3]
                SuspendedUnits= $parts[4]
                Status        = "OK"
            }
            $foundLicenses = $true
        }
        elseif ($line -match "^ERROR\|(.+)$") {
            $errMsg = $Matches[1]
            $results += [PSCustomObject]@{
                Environment   = $row.Environment
                Username      = $user
                TenantId      = $tenantId
                LicenseSKU    = ""
                ActiveUnits   = ""
                ConsumedUnits = ""
                WarningUnits  = ""
                SuspendedUnits= ""
                Status        = "ERROR: $errMsg"
            }
            Write-Host "  ERROR: $errMsg" -ForegroundColor Red
        }
    }

    if ($foundLicenses) {
        $skuCount = ($results | Where-Object { $_.Username -eq $user -and $_.Status -eq "OK" }).Count
        Write-Host "  Found $skuCount license(s)" -ForegroundColor Green
    } elseif (-not $foundLicenses -and -not ($licenseOutput -match "ERROR")) {
        $results += [PSCustomObject]@{
            Environment   = $row.Environment
            Username      = $user
            TenantId      = $tenantId
            LicenseSKU    = ""
            ActiveUnits   = ""
            ConsumedUnits = ""
            WarningUnits  = ""
            SuspendedUnits= ""
            Status        = "NO_LICENSES_FOUND"
        }
        Write-Host "  No licenses found" -ForegroundColor Yellow
    }
}

# =========================================================
# EXPORT RESULTS
# =========================================================
$results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Done! Processed $processed tenants." -ForegroundColor Cyan
Write-Host "Results exported to: $OutputFile" -ForegroundColor Green
Write-Host "Total license rows: $($results.Count)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
