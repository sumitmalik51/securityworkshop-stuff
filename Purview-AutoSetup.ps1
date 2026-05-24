#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Automates Microsoft Purview manual setup steps (DLP, IRM, DSPM roles, Auditing, Analytics).

.DESCRIPTION
    This script automates the following from "Manual Steps - Purview.docx":
    - Connects to Security & Compliance + Exchange Online via certificate auth
    - Enables Organization Customization (required for fresh tenants)
    - Enables Unified Audit Log ingestion
    - Verifies DLP policies are created and enabled
    - Enables advanced classification for endpoint DLP
    - Enables DLP analytics (IsDlpSimulationOptedIn)
    - Assigns ODL user to Insider Risk Management role group
    - Assigns ODL user to Compliance Administrator role group
    - Creates/verifies Purview Content Analyst custom role (DSPM) and assigns ODL user
    - Creates IRM TenantSetting policy if missing and enables IRM analytics
    - Verifies Insider Risk Management policies are enabled with scoring

.NOTES
    Steps that STILL require manual UI interaction:
    - DLP > Overview: Click "Turn on analytics" button (portal consent)
    - DLP/IRM Triage Agent deployment and customization
    - DSPM Posture Agent setup (Exercise 3, Tasks 1-2)
    - DSI setup (Exercise 4, Task 1)
    - Alert trigger scripts (run by Yash)
#>

param(
    [Parameter(Mandatory)]
    [string]$Thumbprint,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [string]$Organization,

    [Parameter(Mandatory)]
    [string]$OdlUser
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "`n[$((Get-Date).ToString('HH:mm:ss'))] $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "  [OK] $Message" -ForegroundColor Green
}

function Write-Skip {
    param([string]$Message)
    Write-Host "  [SKIP] $Message" -ForegroundColor Yellow
}

function Write-Fail {
    param([string]$Message)
    Write-Host "  [FAIL] $Message" -ForegroundColor Red
}

# ============================================================
# STEP 0a: Connect to Exchange Online (for auditing/org cmdlets)
# ============================================================
Write-Step "Connecting to Exchange Online..."

try {
    Import-Module ExchangeOnlineManagement -MinimumVersion 3.0.0
    Connect-ExchangeOnline -CertificateThumbprint $Thumbprint -AppId $ClientId -Organization $Organization -ShowBanner:$false
    Write-Success "Connected EXO to $Organization"
}
catch {
    Write-Fail "EXO Connection failed: $($_.Exception.Message)"
    Write-Host "  Ensure the certificate private key is accessible (check ACLs on the key file)." -ForegroundColor Yellow
    exit 1
}

# ============================================================
# STEP 0b: Enable Organization Customization
# ============================================================
Write-Step "Enabling Organization Customization (if needed)..."

try {
    Enable-OrganizationCustomization -ErrorAction Stop
    Write-Success "Organization Customization enabled. Waiting 60s for propagation..."
    Start-Sleep -Seconds 60
}
catch {
    if ($_.Exception.Message -match "already been enabled|has already been run|already enabled for customization|not required") {
        Write-Skip "Organization Customization is already enabled."
    }
    else {
        Write-Fail "Could not enable Organization Customization: $($_.Exception.Message)"
    }
}

# ============================================================
# STEP 0c: Enable Unified Audit Log
# ============================================================
Write-Step "Enabling Unified Audit Log..."

try {
    $auditConfig = Get-AdminAuditLogConfig
    if ($auditConfig.UnifiedAuditLogIngestionEnabled -eq $true) {
        Write-Skip "Unified Audit Log is already enabled."
    }
    else {
        Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
        Write-Success "Unified Audit Log enabled."
    }
}
catch {
    Write-Fail "Could not enable Audit Log: $($_.Exception.Message)"
    Write-Host "  This may require Enable-OrganizationCustomization to propagate first." -ForegroundColor Yellow
}

# ============================================================
# STEP 0d: Connect to Security & Compliance PowerShell
# ============================================================
Write-Step "Connecting to Security & Compliance PowerShell..."

try {
    Connect-IPPSSession -CertificateThumbprint $Thumbprint -AppId $ClientId -Organization $Organization
    Write-Success "Connected IPPS to $Organization"
}
catch {
    Write-Fail "IPPS Connection failed: $($_.Exception.Message)"
    exit 1
}

# ============================================================
# STEP 1: Verify DLP Policies
# ============================================================
Write-Step "Verifying DLP Policies..."

$dlpPolicies = Get-DlpCompliancePolicy
$enabledCount = ($dlpPolicies | Where-Object { $_.Enabled -eq $true }).Count

if ($enabledCount -gt 0) {
    Write-Success "$enabledCount DLP policies found and enabled:"
    $dlpPolicies | Where-Object { $_.Enabled -eq $true } | ForEach-Object {
        Write-Host "    - $($_.Name)" -ForegroundColor White
    }
}
else {
    Write-Fail "No enabled DLP policies found. Ensure the provisioning script has run."
}

# ============================================================
# STEP 2: Enable Advanced Classification (Endpoint DLP)
# ============================================================
Write-Step "Checking Advanced Classification for Endpoint DLP..."

$policyConfig = Get-PolicyConfig
$endpointSettings = $policyConfig.EndpointDlpGlobalSettings | ConvertFrom-Json -ErrorAction SilentlyContinue

$advClassEnabled = $endpointSettings | Where-Object { $_.Setting -eq "AdvancedClassificationEnabled" -and $_.Value -eq "true" }

if ($advClassEnabled) {
    Write-Skip "Advanced Classification is already enabled."
}
else {
    Write-Host "  Enabling Advanced Classification..." -ForegroundColor White
    try {
        # Build updated settings with AdvancedClassificationEnabled = true
        $settings = @(
            @{ Setting = "AdvancedClassificationEnabled"; Value = "true" }
        )
        Set-PolicyConfig -EndpointDlpGlobalSettings ($settings | ConvertTo-Json -Compress)
        Write-Success "Advanced Classification enabled."
    }
    catch {
        Write-Fail "Could not enable Advanced Classification: $($_.Exception.Message)"
        Write-Host "  You may need to enable this manually in the Purview Portal." -ForegroundColor Yellow
    }
}

# ============================================================
# STEP 2b: Enable DLP Analytics (Backend)
# ============================================================
Write-Step "Enabling DLP Analytics (IsDlpSimulationOptedIn)..."

try {
    $policyConfig = Get-PolicyConfig
    if ($policyConfig.IsDlpSimulationOptedIn -eq $true) {
        Write-Skip "DLP Analytics (IsDlpSimulationOptedIn) is already enabled."
    }
    else {
        Set-PolicyConfig -IsDlpSimulationOptedIn $true
        Write-Success "DLP Analytics backend enabled."
        Write-Host "  NOTE: You still need to click 'Turn on analytics' in the DLP portal." -ForegroundColor Yellow
    }
}
catch {
    Write-Fail "Could not enable DLP Analytics: $($_.Exception.Message)"
}

# ============================================================
# STEP 3: Assign ODL User to Insider Risk Management Role
# ============================================================
Write-Step "Assigning '$OdlUser' to InsiderRiskManagement role group..."

try {
    $members = Get-RoleGroupMember "InsiderRiskManagement"
    $alreadyMember = $members | Where-Object { $_.Name -like "*$($OdlUser.Split('@')[0].Replace('odl_user_',''))*" }

    if ($alreadyMember) {
        Write-Skip "User is already a member of InsiderRiskManagement."
    }
    else {
        Add-RoleGroupMember "InsiderRiskManagement" -Member $OdlUser
        Write-Success "Added to InsiderRiskManagement."
    }
}
catch {
    if ($_.Exception.Message -match "MemberAlreadyExists") {
        Write-Skip "User is already a member of InsiderRiskManagement."
    }
    else {
        Write-Fail "Failed: $($_.Exception.Message)"
    }
}

# ============================================================
# STEP 4: Assign ODL User to Compliance Administrator Role
# ============================================================
Write-Step "Assigning '$OdlUser' to ComplianceAdministrator role group..."

try {
    $members = Get-RoleGroupMember "ComplianceAdministrator"
    $alreadyMember = $members | Where-Object { $_.Name -like "*$($OdlUser.Split('@')[0].Replace('odl_user_',''))*" }

    if ($alreadyMember) {
        Write-Skip "User is already a member of ComplianceAdministrator."
    }
    else {
        Add-RoleGroupMember "ComplianceAdministrator" -Member $OdlUser
        Write-Success "Added to ComplianceAdministrator."
    }
}
catch {
    if ($_.Exception.Message -match "MemberAlreadyExists") {
        Write-Skip "User is already a member of ComplianceAdministrator."
    }
    else {
        Write-Fail "Failed: $($_.Exception.Message)"
    }
}

# ============================================================
# STEP 5: DSPM - Create/Verify Purview Content Analyst Custom Role
# ============================================================
Write-Step "Verifying 'Purview Content Analyst' custom role group (DSPM)..."

$customRoleName = "Purview Content Analyst (Custom)"
$existingRole = Get-RoleGroup | Where-Object { $_.Name -eq $customRoleName }

if (-not $existingRole) {
    Write-Host "  Creating custom role group '$customRoleName'..." -ForegroundColor White
    try {
        # Create with Content Explorer roles for DSPM
        New-RoleGroup -Name $customRoleName -Roles "Data Classification Content Viewer", "Data Classification List Viewer" -Members $OdlUser
        Write-Success "Created '$customRoleName' with ODL user as member."
    }
    catch {
        Write-Fail "Could not create custom role group: $($_.Exception.Message)"
        Write-Host "  You may need to create this manually: Settings > Roles and scope > Add custom role." -ForegroundColor Yellow
    }
}
else {
    Write-Success "Role group '$customRoleName' already exists."
    # Ensure ODL user is a member
    try {
        $members = Get-RoleGroupMember $customRoleName
        $alreadyMember = $members | Where-Object { $_.Name -like "*$($OdlUser.Split('@')[0].Replace('odl_user_',''))*" }

        if ($alreadyMember) {
            Write-Skip "User is already a member of '$customRoleName'."
        }
        else {
            Add-RoleGroupMember $customRoleName -Member $OdlUser
            Write-Success "Added user to '$customRoleName'."
        }
    }
    catch {
        if ($_.Exception.Message -match "MemberAlreadyExists") {
            Write-Skip "User is already a member of '$customRoleName'."
        }
        else {
            Write-Fail "Failed to add member: $($_.Exception.Message)"
        }
    }
}

# ============================================================
# STEP 6: Verify Insider Risk Management Policies
# ============================================================
Write-Step "Verifying Insider Risk Management Policies..."

$irmPolicies = Get-InsiderRiskPolicy | Where-Object { $_.InsiderRiskScenario -ne "TenantSetting" }

if ($irmPolicies.Count -gt 0) {
    $enabledIrm = $irmPolicies | Where-Object { $_.Enabled -eq $true }
    Write-Success "$($enabledIrm.Count)/$($irmPolicies.Count) IRM policies enabled:"
    $enabledIrm | ForEach-Object {
        Write-Host "    - $($_.Name) (Scenario: $($_.InsiderRiskScenario))" -ForegroundColor White
    }
}
else {
    Write-Fail "No IRM policies found. Ensure provisioning scripts have run."
}

# ============================================================
# STEP 7: Ensure IRM TenantSetting & Enable Analytics
# ============================================================
Write-Step "Ensuring IRM TenantSetting policy exists and analytics is enabled..."

$tenantSetting = Get-InsiderRiskPolicy | Where-Object { $_.InsiderRiskScenario -eq "TenantSetting" }

if (-not $tenantSetting) {
    Write-Host "  Creating IRM TenantSetting policy..." -ForegroundColor White
    try {
        New-InsiderRiskPolicy -Name "IRM_TenantSetting" -InsiderRiskScenario TenantSetting
        Write-Success "IRM TenantSetting policy created."
        $tenantSetting = Get-InsiderRiskPolicy -Identity "IRM_TenantSetting"
    }
    catch {
        Write-Fail "Could not create IRM TenantSetting: $($_.Exception.Message)"
    }
}
else {
    Write-Skip "IRM TenantSetting policy already exists."
}

if ($tenantSetting) {
    try {
        Set-InsiderRiskPolicy -Identity $tenantSetting.Name -TurnOnAnalytics $true
        Write-Success "IRM Analytics (TurnOnAnalytics) enabled."
    }
    catch {
        if ($_.Exception.Message -match "already enabled") {
            Write-Skip "IRM Analytics is already enabled."
        }
        else {
            Write-Fail "Could not enable IRM Analytics: $($_.Exception.Message)"
        }
    }
}

# ============================================================
# SUMMARY
# ============================================================
Write-Host "`n" -NoNewline
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host " AUTOMATION COMPLETE - Summary for $Organization" -ForegroundColor Magenta
Write-Host "============================================================" -ForegroundColor Magenta
Write-Host ""
Write-Host " Automated steps completed above." -ForegroundColor White
Write-Host ""
Write-Host " MANUAL STEPS REMAINING (Purview Portal UI):" -ForegroundColor Yellow
Write-Host "  1. DLP > Overview: Click 'Turn on analytics' button in recommendation" -ForegroundColor Yellow
Write-Host "  2. DLP > Agents: Deploy Triage Agent (Automatic trigger, 30 days)" -ForegroundColor Yellow
Write-Host "  3. DLP > Triage Agent: Customize scope (all policies) + custom instruction:" -ForegroundColor Yellow
Write-Host "     'Focus on alerts with content that is tax or finance related" -ForegroundColor Gray
Write-Host "      and contains more than five credit card numbers or SSNs'" -ForegroundColor Gray
Write-Host "  4. IRM > Agents: Deploy Triage Agent (Manual, one alert at a time)" -ForegroundColor Yellow
Write-Host "  5. IRM > Triage Agent: Customize scope (select required policies)" -ForegroundColor Yellow
Write-Host "  6. DSPM: Enable & configure Posture Agent (Lab Exercise 3, Tasks 1-2)" -ForegroundColor Yellow
Write-Host "  7. DSI: Complete Lab Exercise 4, Task 1" -ForegroundColor Yellow
Write-Host ""
Write-Host "============================================================" -ForegroundColor Magenta

# Disconnect sessions
Write-Step "Disconnecting sessions..."
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
Write-Success "Done."
