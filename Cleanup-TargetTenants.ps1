<#
.SYNOPSIS
    Cleanup specific orphaned tenants - removes all locks and deletes all RGs.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $ExcelFile = 'C:\certs\spns.xlsx'
)

$ErrorActionPreference = 'Continue'

$targetOrgs = @(
    'otuwacne104045','otuwacne104046','otuwacne104047','otuwacne104054','otuwacne104055',
    'otuwacne104056','otuwacne104057','otuwacne104058','otuwacne104060','otuwacne104061',
    'otuwacne104062','otuwacne104063','otuwacne104064','otuwacne104041','otuwacne104042',
    'otuwacne104043','otuwacne104044','otuwacne104052','otuwacne104051','otuwacne104053',
    'otuwacne104048','otuwacne104065','otuwacne104050','otuwacne103923','otuwacne103926',
    'otuwacne103982','otuwacne103991','otuwacne104127','otuwacne104092','otuwacne104080',
    'otuwacne104099','otuwacne104131','otuwacne104132','otuwacne104133','otuwacne104134',
    'otuwacne104135','otuwacne104136','otuwacne104137','otuwacne104138','otuwacne104139',
    'otuwacne104140','otuwacne103892','otuwacne103897','otuwacne103900','otuwacne103902',
    'otuwacne103906','otuwacne103930','otuwacne103937','otuwacne103946','otuwacne103947',
    'otuwacne103950','otuwacne103956','otuwacne103957','otuwacne103958','otuwacne103960',
    'otuwacne103971','otuwacne103974','otuwacne103977','otuwacne103989','otuwacne103999',
    'otuwacne104009','otuwacne104119'
)

# ── load matching accounts from all batches ──
$accounts = @()
foreach ($batch in @('Batch1','Batch2','Batch3','Batch4','Batch5')) {
    $data = Import-Excel -Path $ExcelFile -WorksheetName $batch
    foreach ($row in $data) {
        foreach ($t in $targetOrgs) {
            if ($row.orgname -like "*$t*") {
                $accounts += $row
                break
            }
        }
    }
}

$total = $accounts.Count
Write-Host "Found $total target tenants to clean up" -ForegroundColor Cyan
Write-Host ""

$global:processed = 0
$global:failed    = 0
$global:cleaned   = 0
$results = @()

for ($i = 0; $i -lt $total; $i++) {
    $acct = $accounts[$i]
    $idx  = $i + 1
    $short = ($acct.odluser -split '@')[0]
    $org   = ($acct.orgname -replace '\.onmicrosoft\.com','')

    Write-Host "`n=== [$idx/$total] $short ($org) ===" -ForegroundColor Cyan

    if (-not $acct.SubscriptionId) {
        Write-Host "  No SubscriptionId - skipping" -ForegroundColor DarkGray
        $results += [pscustomobject]@{ Index=$idx; User=$short; Org=$org; Status="NO_SUB"; RGsDeleted=0 }
        $global:processed++
        continue
    }

    # ── authenticate ──
    try {
        $secSecret = ConvertTo-SecureString $acct.AppSecret -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($acct.AppId, $secSecret)
        Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $acct.TenantId -ErrorAction Stop | Out-Null
        Select-AzSubscription -SubscriptionId $acct.SubscriptionId -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "  AUTH FAILED: $_" -ForegroundColor Red
        $global:failed++
        $results += [pscustomobject]@{ Index=$idx; User=$short; Org=$org; Status="AUTH_FAILED"; RGsDeleted=0 }
        continue
    }

    # ── get all RGs ──
    $rgs = Get-AzResourceGroup -ErrorAction SilentlyContinue
    if (-not $rgs) {
        Write-Host "  No RGs found - already clean" -ForegroundColor DarkGray
        $results += [pscustomobject]@{ Index=$idx; User=$short; Org=$org; Status="ALREADY_CLEAN"; RGsDeleted=0 }
        $global:processed++
        continue
    }

    $rgCount = 0
    foreach ($rg in $rgs) {
        $rgName = $rg.ResourceGroupName
        Write-Host "    Deleting RG: $rgName" -ForegroundColor Yellow

        # Remove all locks
        $locks = Get-AzResourceLock -ResourceGroupName $rgName -ErrorAction SilentlyContinue
        foreach ($l in $locks) {
            Write-Host "      Removing lock '$($l.Name)'" -ForegroundColor DarkYellow
            Remove-AzResourceLock -LockId $l.LockId -Force -ErrorAction SilentlyContinue | Out-Null
        }

        # Delete RG
        if ($PSCmdlet.ShouldProcess($rgName, "Delete resource group")) {
            try {
                Remove-AzResourceGroup -Name $rgName -Force -ErrorAction Stop | Out-Null
                Write-Host "      Deleted $rgName" -ForegroundColor Green
                $rgCount++
            } catch {
                Write-Host "      Failed: $_" -ForegroundColor Red
            }
        }
    }

    Write-Host "  CLEANED - $rgCount/$($rgs.Count) RGs deleted" -ForegroundColor Magenta
    $results += [pscustomobject]@{ Index=$idx; User=$short; Org=$org; Status="CLEANED"; RGsDeleted=$rgCount }
    $global:cleaned++
    $global:processed++
}

# ── summary ──
Write-Host "`n======================================="
Write-Host "SUMMARY"
Write-Host "======================================="
Write-Host "Total:   $total"
Write-Host "Cleaned: $($global:cleaned)"
Write-Host "Failed:  $($global:failed)"
Write-Host ""
$results | Format-Table -AutoSize

$csvPath = "C:\certs\Orphaned-Cleanup-Results.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Force
Write-Host "`nResults saved to $csvPath"
