<#
.SYNOPSIS
    Removes ALL locks (RG-level and resource-level) from ALL accounts across all batches.
    Does NOT delete any resources or resource groups — locks only.

.EXAMPLE
    .\Remove-AllLocks.ps1
    .\Remove-AllLocks.ps1 -Batches Batch4,Batch5
    .\Remove-AllLocks.ps1 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]   $ExcelFile = 'C:\certs\spns.xlsx',
    [string[]] $Batches   = @('Batch1','Batch2','Batch3','Batch4','Batch5')
)

$ErrorActionPreference = 'Continue'
Import-Module ImportExcel -ErrorAction Stop

# ── load all accounts ──
$accounts = @()
foreach ($batch in $Batches) {
    $data = Import-Excel -Path $ExcelFile -WorksheetName $batch
    foreach ($row in $data) {
        $row | Add-Member -NotePropertyName 'Batch' -NotePropertyValue $batch -Force
        $accounts += $row
    }
}

$total = $accounts.Count
Write-Host "Loaded $total accounts from batches: $($Batches -join ', ')" -ForegroundColor Cyan
Write-Host ""

$processed  = 0
$skipped    = 0
$authFailed = 0
$noLocks    = 0
$lockTotal  = 0
$results    = @()

foreach ($acct in $accounts) {
    $processed++
    $sub   = $acct.SubscriptionId
    $tid   = $acct.TenantId
    $appId = $acct.AppId
    $sec   = $acct.AppSecret
    $org   = $acct.orgname
    $user  = ($acct.odluser -split '@')[0]
    $batch = $acct.Batch

    # Skip accounts without credentials
    if (-not $sub -or -not $appId -or -not $sec) {
        Write-Host "[$processed/$total] $user ($org) - SKIP (no credentials)" -ForegroundColor DarkGray
        $skipped++
        continue
    }

    # Authenticate
    try {
        $cred = New-Object PSCredential($appId, (ConvertTo-SecureString $sec -AsPlainText -Force))
        Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $tid -Subscription $sub -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "[$processed/$total] $user ($org) - AUTH_FAILED" -ForegroundColor Red
        $authFailed++
        continue
    }

    # Get all locks in the subscription
    try {
        $locks = Get-AzResourceLock -ErrorAction Stop
    } catch {
        Write-Host "[$processed/$total] $user ($org) - ERROR listing locks: $_" -ForegroundColor Red
        $results += [pscustomobject]@{ Index=$processed; User=$user; Org=$org; Batch=$batch; Status="ERROR"; LocksRemoved=0; Details="$_" }
        continue
    }

    if ($locks.Count -eq 0) {
        Write-Host "[$processed/$total] $user ($org) - no locks" -ForegroundColor DarkGray
        $noLocks++
        continue
    }

    Write-Host "[$processed/$total] $user ($org) - $($locks.Count) lock(s)" -ForegroundColor White -NoNewline
    $removed = 0

    foreach ($lock in $locks) {
        if ($PSCmdlet.ShouldProcess("$($lock.Name) on $($lock.ResourceGroupName)", "Remove lock")) {
            try {
                Remove-AzResourceLock -LockId $lock.LockId -Force -ErrorAction Stop | Out-Null
                $removed++
            } catch {
                # Ignore already-removed locks
            }
        }
    }

    $lockTotal += $removed
    Write-Host " -> $removed removed" -ForegroundColor Green
    $results += [pscustomobject]@{ Index=$processed; User=$user; Org=$org; Batch=$batch; Status="CLEANED"; LocksRemoved=$removed; Details="" }
}

# ── summary ──
Write-Host "`n=======================================" -ForegroundColor Cyan
Write-Host "LOCK REMOVAL SUMMARY" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "Total accounts:    $total"
Write-Host "Processed:         $processed"
Write-Host "Skipped (no cred): $skipped"
Write-Host "Auth failed:       $authFailed"
Write-Host "No locks found:    $noLocks"
Write-Host "Total locks removed: $lockTotal" -ForegroundColor Green

if ($results.Count -gt 0) {
    $csv = 'C:\certs\LockRemoval-Results.csv'
    $results | Export-Csv -Path $csv -NoTypeInformation -Force
    Write-Host "Results saved to $csv"
}
