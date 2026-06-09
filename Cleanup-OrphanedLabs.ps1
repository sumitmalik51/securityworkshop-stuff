<#
.SYNOPSIS
    Checks if the ODL user exists in the Entra tenant; if not, deletes all resource groups in the subscription.

.DESCRIPTION
    For each tenant/subscription in the specified batches from spns.xlsx:
      1. Authenticates via SPN (AppId/AppSecret)
      2. Checks if the odluser exists in Entra ID (Azure AD) using Microsoft Graph
      3. If the user does NOT exist, removes all resource locks and deletes every resource group

.EXAMPLE
    .\Cleanup-OrphanedLabs.ps1 -Batches Batch1,Batch2
    .\Cleanup-OrphanedLabs.ps1 -Batches Batch5 -WhatIf

.NOTES
    Requires: Az.Accounts, Az.Resources, ImportExcel
    The SPN must have User.Read.All (or Directory.Read.All) Graph permission to look up users,
    and Contributor + User Access Administrator on the subscription to delete RGs.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]] $Batches   = @('Batch1','Batch2','Batch3','Batch4','Batch5'),
    [string]   $ExcelFile = 'C:\certs\spns.xlsx'
)

$ErrorActionPreference = 'Continue'

# ── helpers ──────────────────────────────────────────────
function Remove-AllLocksInRG {
    param([string]$ResourceGroupName)
    $locks = Get-AzResourceLock -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue
    foreach ($l in $locks) {
        Write-Host "      Removing lock '$($l.Name)'" -ForegroundColor Yellow
        Remove-AzResourceLock -LockId $l.LockId -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

function Test-UserExistsInTenant {
    param(
        [string]$UserPrincipalName,
        [string]$TenantId,
        [string]$AppId,
        [System.Security.SecureString]$AppSecret
    )
    # Get an access token for Microsoft Graph using the SPN credentials
    try {
        $plainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($AppSecret))
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $AppId
            client_secret = $plainSecret
            scope         = 'https://graph.microsoft.com/.default'
        }
        $tokenResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body $body -ErrorAction Stop

        $headers = @{ Authorization = "Bearer $($tokenResponse.access_token)" }

        # Try to find the user by UPN
        $filter = "userPrincipalName eq '$UserPrincipalName'"
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=id,userPrincipalName&`$top=1"
        $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
        return ($resp.value.Count -gt 0)
    }
    catch {
        Write-Host "      Graph lookup failed: $_" -ForegroundColor DarkYellow
        # If we can't query Graph (no permissions), return $true to be safe (skip cleanup)
        return $true
    }
}

# ── load accounts ────────────────────────────────────────
$accounts = @()
foreach ($batch in $Batches) {
    $accounts += Import-Excel -Path $ExcelFile -WorksheetName $batch
}
$total = $accounts.Count
Write-Host "Loaded $total accounts from batches: $($Batches -join ', ')" -ForegroundColor Cyan
Write-Host ""

# ── process each tenant ──────────────────────────────────
$global:processed = 0
$global:failed    = 0
$global:cleaned   = 0
$global:skipped   = 0
$results = @()

for ($i = 0; $i -lt $total; $i++) {
    $acct = $accounts[$i]
    $idx  = $i + 1
    $short = ($acct.odluser -split '@')[0]

    Write-Host "`n=== [$idx/$total] $short ===" -ForegroundColor Cyan

    # ── skip if no SubscriptionId ──
    if (-not $acct.SubscriptionId) {
        Write-Host "  No SubscriptionId - skipping (unprovisioned)" -ForegroundColor DarkGray
        $results += [pscustomobject]@{ Index=$idx; User=$short; Status="NO_SUB"; UserExists="N/A"; RGsDeleted=0 }
        $global:processed++
        continue
    }

    # ── authenticate as SPN ──
    try {
        $secSecret = ConvertTo-SecureString $acct.AppSecret -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($acct.AppId, $secSecret)
        Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $acct.TenantId -ErrorAction Stop | Out-Null
        Select-AzSubscription -SubscriptionId $acct.SubscriptionId -ErrorAction Stop | Out-Null
        Write-Host "  Authenticated to subscription $($acct.SubscriptionId)" -ForegroundColor Green
    } catch {
        Write-Host "  FAILED to authenticate: $_" -ForegroundColor Red
        $global:failed++
        $results += [pscustomobject]@{ Index=$idx; User=$short; Status="AUTH_FAILED"; UserExists="N/A"; RGsDeleted=0 }
        continue
    }

    # ── check if odluser exists in Entra ──
    $userExists = Test-UserExistsInTenant -UserPrincipalName $acct.odluser `
                                          -TenantId $acct.TenantId `
                                          -AppId $acct.AppId `
                                          -AppSecret $secSecret

    if ($userExists) {
        Write-Host "  User '$($acct.odluser)' EXISTS in tenant - skipping cleanup" -ForegroundColor Green
        $results += [pscustomobject]@{ Index=$idx; User=$short; Status="USER_EXISTS"; UserExists="Yes"; RGsDeleted=0 }
        $global:skipped++
        $global:processed++
        continue
    }

    Write-Host "  User '$($acct.odluser)' NOT FOUND in tenant - cleaning up subscription" -ForegroundColor Red

    # ── get all resource groups ──
    $rgs = Get-AzResourceGroup -ErrorAction SilentlyContinue
    if (-not $rgs) {
        Write-Host "  No resource groups found" -ForegroundColor DarkGray
        $results += [pscustomobject]@{ Index=$idx; User=$short; Status="NO_USER_NO_RGs"; UserExists="No"; RGsDeleted=0 }
        $global:cleaned++
        $global:processed++
        continue
    }

    $rgCount = 0
    foreach ($rg in $rgs) {
        $rgName = $rg.ResourceGroupName
        Write-Host "    Deleting RG: $rgName" -ForegroundColor Yellow

        # Remove all locks first
        Remove-AllLocksInRG -ResourceGroupName $rgName

        # Delete the resource group
        if ($PSCmdlet.ShouldProcess($rgName, "Delete resource group")) {
            try {
                Remove-AzResourceGroup -Name $rgName -Force -ErrorAction Stop | Out-Null
                Write-Host "      Deleted $rgName" -ForegroundColor Green
                $rgCount++
            } catch {
                Write-Host "      Failed to delete $rgName : $_" -ForegroundColor Red
            }
        }
    }

    Write-Host "  CLEANED - $rgCount RGs deleted" -ForegroundColor Magenta
    $results += [pscustomobject]@{ Index=$idx; User=$short; Status="CLEANED"; UserExists="No"; RGsDeleted=$rgCount }
    $global:cleaned++
    $global:processed++
}

# ── summary ──────────────────────────────────────────────
$batchLabel = ($Batches -join '-')

Write-Host "`n======================================="
Write-Host "SUMMARY"
Write-Host "======================================="
Write-Host "Total:     $total"
Write-Host "Processed: $($global:processed)"
Write-Host "Skipped (user exists): $($global:skipped)"
Write-Host "Cleaned (user gone):   $($global:cleaned)"
Write-Host "Failed:    $($global:failed)"

$failColor = if ($global:failed -gt 0) { 'Red' } else { 'Green' }
Write-Host "Failed:    $($global:failed)" -ForegroundColor $failColor
Write-Host ""

$results | Format-Table -AutoSize

# ── export CSV ───────────────────────────────────────────
$csvPath = "C:\certs\Orphaned-Lab-Cleanup-$batchLabel.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Force
Write-Host "`nResults saved to $csvPath"

exit $(if ($global:failed -gt 0) { 1 } else { 0 })
