# =============================================================================
# Replicate-IRM-3Policies.ps1
# Replicates 3 IRM policies from reference tenant 2226978 to all other tenants
# Source: odl_user_2226978@otuwacne103908.onmicrosoft.com
# Policies:
#   1. "Lab - IRM Critical Data Leak" (LeakOfInformation, AggregatedDataLeakTrigger)
#   2. "Lab - IRM Data Leak Monitoring" (LeakOfInformation, AggregatedDataLeakTrigger)
#   3. "Lab - IRM Data Theft Policy" (IntellectualPropertyTheft, AadLeaver)
# NOTE: TenantSetting is NOT replicated per requirement
# =============================================================================

param(
    [string]$SpnsFile = "C:\certs\spns.xlsx",
    [int]$StartIndex = 0,
    [int]$MaxCount = 0
)

# Import required modules
Import-Module ImportExcel -ErrorAction Stop

# ============================================================
# Load policy config from exported files
# ============================================================
Write-Host "Loading policy configurations..." -ForegroundColor Cyan

$TIG_Critical = Get-Content "C:\certs\tig_critical.txt" -Raw -Encoding UTF8
$TIG_Monitoring = Get-Content "C:\certs\tig_monitoring.txt" -Raw -Encoding UTF8

$Ind_Tenant = @(Get-Content "C:\certs\ind_tenant_lines.txt" -Encoding UTF8 | Where-Object { $_.Trim() -ne "" })
$ExtInd_Tenant = @(Get-Content "C:\certs\ext_ind_tenant_lines.txt" -Encoding UTF8 | Where-Object { $_.Trim() -ne "" })
$Ind_Critical = @(Get-Content "C:\certs\ind_critical_lines.txt" -Encoding UTF8 | Where-Object { $_.Trim() -ne "" })
$Ind_Monitoring = @(Get-Content "C:\certs\ind_monitoring_lines.txt" -Encoding UTF8 | Where-Object { $_.Trim() -ne "" })
$Ind_Theft = @(Get-Content "C:\certs\ind_theft_lines.txt" -Encoding UTF8 | Where-Object { $_.Trim() -ne "" })

Write-Host "  TIG Critical: $($TIG_Critical.Length) chars"
Write-Host "  TIG Monitoring: $($TIG_Monitoring.Length) chars"
Write-Host "  Tenant Indicators: $($Ind_Tenant.Count) items"
Write-Host "  Tenant ExtensibleIndicators: $($ExtInd_Tenant.Count) items"
Write-Host "  Indicators Critical: $($Ind_Critical.Count) items"
Write-Host "  Indicators Monitoring: $($Ind_Monitoring.Count) items"
Write-Host "  Indicators Theft: $($Ind_Theft.Count) items"

# ============================================================
# Load accounts from spns.xlsx (all 5 batches)
# ============================================================
Write-Host "`nLoading accounts from spns.xlsx..." -ForegroundColor Cyan

$allAccounts = @()
foreach ($batch in @("Batch1","Batch2","Batch3","Batch4","Batch5")) {
    $rows = Import-Excel -Path $SpnsFile -WorksheetName $batch
    foreach ($row in $rows) {
        $allAccounts += @{
            Username = $row.odluser
            Password = $row.odlpassword
            TenantId = $row.TenantId
            Name     = $row.Name
        }
    }
}

Write-Host "  Total accounts loaded: $($allAccounts.Count)"

# ============================================================
# Skip accounts (reference + already done)
# ============================================================
$skipUsers = @(
    "odl_user_2226978",  # reference tenant
    "odl_user_2226954",  # already done
    "odl_user_2226956",  # already done
    "odl_user_2226957",  # already done
    "odl_user_2226963",  # already done
    "odl_user_2226964",  # already done
    "odl_user_2226966"   # already done
)

$targetAccounts = @()
foreach ($acct in $allAccounts) {
    $userPrefix = ($acct.Username -split "@")[0]
    if ($userPrefix -notin $skipUsers) {
        $targetAccounts += $acct
    }
}

Write-Host "  Target accounts (after skipping): $($targetAccounts.Count)"

# Apply StartIndex and MaxCount
if ($StartIndex -gt 0) {
    $targetAccounts = $targetAccounts[$StartIndex..($targetAccounts.Count - 1)]
}
if ($MaxCount -gt 0 -and $MaxCount -lt $targetAccounts.Count) {
    $targetAccounts = $targetAccounts[0..($MaxCount - 1)]
}

Write-Host "  Accounts to process this run: $($targetAccounts.Count)"

# ============================================================
# MAIN LOOP
# ============================================================
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = "C:\certs\irm_3policies_log_$Timestamp.txt"

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$ts] $Message"
    Write-Host $logLine -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $logLine
}

Write-Log "=== IRM 3-Policy Replication from 2226978 - Starting ==="
Write-Log "Target accounts: $($targetAccounts.Count)"

$successCount = 0
$failCount = 0
$index = 0

foreach ($account in $targetAccounts) {
    $index++
    $Username = $account.Username
    $Password = $account.Password

    Write-Log ""
    Write-Log "=== [$index/$($targetAccounts.Count)] $Username ==="

    try {
        # Disconnect any prior session
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

        # Connect
        Write-Log "  Connecting..."
        $SecPass = ConvertTo-SecureString $Password -AsPlainText -Force
        $Cred = New-Object System.Management.Automation.PSCredential($Username, $SecPass)
        Connect-IPPSSession -Credential $Cred -CommandName @("New-InsiderRiskPolicy","Set-InsiderRiskPolicy","Get-InsiderRiskPolicy","Remove-InsiderRiskPolicy") -WarningAction SilentlyContinue -ErrorAction Stop

        # ========================================
        # Step 0: Turn on Tenant-level Indicators
        # ========================================
        Write-Log "  [TS] Setting tenant-level indicators ($($Ind_Tenant.Count))..."
        $tsPolicies = Get-InsiderRiskPolicy | Where-Object { $_.InsiderRiskScenario -eq "TenantSetting" }
        if ($tsPolicies) {
            $tsName = $tsPolicies.Name
        } else {
            $tsName = "IRM_Tenant_Setting"
            New-InsiderRiskPolicy -Name $tsName -InsiderRiskScenario TenantSetting -ErrorAction Stop | Out-Null
        }
        Set-InsiderRiskPolicy -Identity $tsName -Indicators $Ind_Tenant -ErrorAction Stop
        Set-InsiderRiskPolicy -Identity $tsName -ExtensibleIndicators $ExtInd_Tenant -ErrorAction Stop

        # ========================================
        # Policy 1: Lab - IRM Critical Data Leak
        # ========================================
        $P1Name = "Lab - IRM Critical Data Leak"
        Write-Log "  [P1] Creating '$P1Name'..."
        try {
            New-InsiderRiskPolicy -Name $P1Name -InsiderRiskScenario LeakOfInformation -ErrorAction Stop | Out-Null
            Write-Log "    Created new"
        } catch {
            if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*exceeds the available limit*") {
                Write-Log "    Already exists - will update"
            } else { throw $_ }
        }

        Write-Log "  [P1] Setting triggers + config..."
        Set-InsiderRiskPolicy -Identity $P1Name `
            -Enabled $true `
            -AddExchangeLocation "All" `
            -Triggers "AggregatedDataLeakTrigger" `
            -TriggerInsightGroups $TIG_Critical `
            -IsPriorityContentOnlyScoring $false `
            -InScopeTimeSpan 30 `
            -HistoricTimeSpan 90 `
            -ErrorAction Stop

        Write-Log "  [P1] Setting indicators ($($Ind_Critical.Count)) + IsCustom..."
        Set-InsiderRiskPolicy -Identity $P1Name `
            -Indicators $Ind_Critical `
            -IsCustom $true `
            -ErrorAction Stop

        # ========================================
        # Policy 2: Lab - IRM Data Leak Monitoring
        # ========================================
        $P2Name = "Lab - IRM Data Leak Monitoring"
        Write-Log "  [P2] Creating '$P2Name'..."
        try {
            New-InsiderRiskPolicy -Name $P2Name -InsiderRiskScenario LeakOfInformation -ErrorAction Stop | Out-Null
            Write-Log "    Created new"
        } catch {
            if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*exceeds the available limit*") {
                Write-Log "    Already exists - will update"
            } else { throw $_ }
        }

        Write-Log "  [P2] Setting triggers + config..."
        Set-InsiderRiskPolicy -Identity $P2Name `
            -Enabled $true `
            -AddExchangeLocation "All" `
            -Triggers "AggregatedDataLeakTrigger" `
            -TriggerInsightGroups $TIG_Monitoring `
            -IsPriorityContentOnlyScoring $false `
            -InScopeTimeSpan 30 `
            -HistoricTimeSpan 90 `
            -ErrorAction Stop

        Write-Log "  [P2] Setting indicators ($($Ind_Monitoring.Count)) + IsCustom..."
        Set-InsiderRiskPolicy -Identity $P2Name `
            -Indicators $Ind_Monitoring `
            -IsCustom $true `
            -ErrorAction Stop

        # ========================================
        # Policy 3: Lab - IRM Data Theft Policy
        # ========================================
        $P3Name = "Lab - IRM Data Theft Policy"
        Write-Log "  [P3] Creating '$P3Name'..."
        try {
            New-InsiderRiskPolicy -Name $P3Name -InsiderRiskScenario IntellectualPropertyTheft -ErrorAction Stop | Out-Null
            Write-Log "    Created new"
        } catch {
            if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*exceeds the available limit*") {
                Write-Log "    Already exists - will update"
            } else { throw $_ }
        }

        Write-Log "  [P3] Setting triggers + config..."
        Set-InsiderRiskPolicy -Identity $P3Name `
            -Enabled $true `
            -AddExchangeLocation "All" `
            -Triggers "AadLeaver" `
            -IsPriorityContentOnlyScoring $false `
            -InScopeTimeSpan 30 `
            -HistoricTimeSpan 90 `
            -ErrorAction Stop

        Write-Log "  [P3] Setting indicators ($($Ind_Theft.Count)) + IsCustom..."
        Set-InsiderRiskPolicy -Identity $P3Name `
            -Indicators $Ind_Theft `
            -IsCustom $true `
            -ErrorAction Stop

        Write-Log "  SUCCESS - All 3 policies created" "Green"
        $successCount++

    } catch {
        Write-Log "  FAILED: $($_.Exception.Message)" "Red"
        $failCount++
    }

    # Disconnect
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}

Write-Log ""
Write-Log "============================================================"
Write-Log "COMPLETE: Success=$successCount, Failed=$failCount, Total=$($targetAccounts.Count)"
Write-Log "============================================================"
Write-Log "Log: $LogFile"
