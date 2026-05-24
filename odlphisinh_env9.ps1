

$inputFile = "C:\certs\66503-User-Detail-Report-separate.xlsx"

# =========================================================
# IMPORT EXCEL
# =========================================================

$users = Import-Excel `
    -Path $inputFile `
    -WorksheetName "Environment Detail 9"
    # =========================================================
# START ALL VMS + RUN IDLE TRACKER CLEANUP SCRIPT
# =========================================================


# =========================================================
# RESULTS
# =========================================================

$results = @()









foreach ($row in $users) {

    try {

        $user     = $row.username
        $pass     = $row.password
        $tenantId = $row.'tenant id'

        Write-Host ""
        Write-Host "================================="
        Write-Host "Processing: $user"
        Write-Host "================================="

        $domain   = $user.Split("@")[1]
        $orgName  = $domain.Replace(".onmicrosoft.com","")

        # Reset counters
        $sent = 0

$domain   = $user.Split("@")[1]
$orgName  = $domain.Replace(".onmicrosoft.com","")

Write-Host "Phishing Triage - Complete Setup"
Write-Host "Tenant: $domain"
Write-Host "User: $user"

# ================================================================
# PHASE 1: EXO — Shared Mailboxes + Permissions + Report Policy
# ================================================================
Write-Host ""
Write-Host "PHASE 1: Exchange Online Setup"

# Must run in clean subprocess to avoid DLL/broker conflicts
$exoResult = powershell.exe -NoProfile -ExecutionPolicy Bypass -Command {
    param($user, $pass, $domain)

    $env:AZURE_ENABLE_WAM = "false"
    [System.Environment]::SetEnvironmentVariable("Broker_Enabled","false","Process")

    # Auto-install EXO module if missing
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "Installing ExchangeOnlineManagement..."
        Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber
    }

    $sp = ConvertTo-SecureString $pass -AsPlainText -Force
    $cr = New-Object System.Management.Automation.PSCredential($user, $sp)

    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -Credential $cr -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected to Exchange Online"

    # --- Step 1: Create shared mailboxes ---
    Write-Host ""
    Write-Host "Step 1: Creating shared mailboxes..."

    $itSec = Get-Mailbox -Identity "itsecurity@$domain" -ErrorAction SilentlyContinue
    if ($itSec) {
        Write-Host "  itsecurity@ already exists"
    } else {
        try {
            New-Mailbox -Shared -Name "IT Security Team" -Alias "itsecurity" -PrimarySmtpAddress "itsecurity@$domain" -ErrorAction Stop | Out-Null
            Write-Host "  Created itsecurity@$domain"
        } catch { Write-Host "  itsecurity@ error: $_" }
    }

    $finDept = Get-Mailbox -Identity "finance-dept@$domain" -ErrorAction SilentlyContinue
    if ($finDept) {
        Write-Host "  finance-dept@ already exists"
    } else {
        try {
            New-Mailbox -Shared -Name "Finance Department" -Alias "finance-dept" -PrimarySmtpAddress "finance-dept@$domain" -ErrorAction Stop | Out-Null
            Write-Host "  Created finance-dept@$domain"
        } catch { Write-Host "  finance-dept@ error: $_" }
    }

    # --- Step 2: Grant permissions ---
    Write-Host ""
    Write-Host "Step 2: Granting FullAccess + SendAs permissions..."

    foreach ($mbx in @("itsecurity@$domain", "finance-dept@$domain")) {
        try {
            Add-MailboxPermission -Identity $mbx -User $user -AccessRights FullAccess -AutoMapping $false -ErrorAction Stop | Out-Null
            Write-Host "  FullAccess on $mbx"
        } catch {
            if ($_.Exception.Message -match "already exists") { Write-Host "  FullAccess on $mbx (already granted)" }
            else { Write-Host "  FullAccess $mbx error: $_" }
        }
        try {
            Add-RecipientPermission -Identity $mbx -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
            Write-Host "  SendAs on $mbx"
        } catch {
            if ($_.Exception.Message -match "already present") { Write-Host "  SendAs on $mbx (already granted)" }
            else { Write-Host "  SendAs $mbx error: $_" }
        }
    }

    # --- Step 3: Configure Report Submission Policy ---
    Write-Host ""
    Write-Host "Step 3: Configuring Report Submission Policy..."

    try {
        $existingPolicy = Get-ReportSubmissionPolicy -ErrorAction SilentlyContinue
        if ($existingPolicy) {
            Set-ReportSubmissionPolicy -Identity DefaultReportSubmissionPolicy `
                -EnableReportToMicrosoft $true `
                -ReportJunkToCustomizedAddress $false `
                -ReportNotJunkToCustomizedAddress $false `
                -ReportPhishToCustomizedAddress $false `
                -EnableThirdPartyAddress $false `
                -ReportChatMessageEnabled $false `
                -PreSubmitMessageEnabled $true `
                -PostSubmitMessageEnabled $true `
                -EnableUserEmailNotification $true `
                -ErrorAction Stop
            Write-Host "  Updated Report Submission Policy"
        } else {
            New-ReportSubmissionPolicy `
                -EnableReportToMicrosoft $true `
                -ReportJunkToCustomizedAddress $false `
                -ReportNotJunkToCustomizedAddress $false `
                -ReportPhishToCustomizedAddress $false `
                -ErrorAction Stop
            Write-Host "  Created Report Submission Policy"
        }
    } catch {
        if ($_.Exception.Message -match "no settings.*modified") { Write-Host "  Policy already configured" }
        else { Write-Host "  Policy: $_" }
    }

    # --- Step 4: Create Report Submission Rule ---
    Write-Host ""
    Write-Host "Step 4: Creating Report Submission Rule..."

    try {
        $existingRule = Get-ReportSubmissionRule -ErrorAction SilentlyContinue
        if ($existingRule) {
            Write-Host "  Rule already exists: $($existingRule.Name) | State: $($existingRule.State)"
            if ($existingRule.State -ne "Enabled") {
                Enable-ReportSubmissionRule -Identity $existingRule.Name -ErrorAction Stop
                Write-Host "  Rule enabled"
            }
        } else {
            # Try without -SentTo first (works on some EXO versions)
            try {
                New-ReportSubmissionRule -Name "DefaultReportSubmissionRule" `
                    -ReportSubmissionPolicy "DefaultReportSubmissionPolicy" `
                    -ErrorAction Stop
                Write-Host "  Rule created"
            } catch {
                # Try with -SentTo (required on newer EXO)
                try {
                    New-ReportSubmissionRule -Name "DefaultReportSubmissionRule" `
                        -ReportSubmissionPolicy "DefaultReportSubmissionPolicy" `
                        -SentTo $user `
                        -ErrorAction Stop
                    Write-Host "  Rule created (with SentTo)"
                } catch {
                    Write-Host "  Rule creation failed: $_"
                    Write-Host "  NOTE: Report button may still work via Defender portal built-in reporting"
                }
            }
        }
    } catch {
        Write-Host "  Rule check: $_"
    }

    # --- Step 5: Deploy Report Message Add-in ---
    Write-Host ""
    Write-Host "Step 5: Deploying Report Message Add-in..."

    try {
        $addins = Get-App -OrganizationApp -ErrorAction SilentlyContinue
        $reportAddin = $addins | Where-Object { $_.DisplayName -match "Report Message|Report Phishing" }
        if ($reportAddin) {
            Write-Host "  Add-in already deployed: $($reportAddin.DisplayName)"
            Set-App -Identity $reportAddin.AppId -OrganizationApp -DefaultStateForUser Enabled -ProvidedTo Everyone -ErrorAction SilentlyContinue
            Write-Host "  Enabled for all users"
        } else {
            # Method 1: MarketplaceAssetId
            try {
                New-App -OrganizationApp -MarketplaceAssetId "WA104381180" `
                    -MarketplaceQueryMarket "en-US" `
                    -DefaultStateForUser Enabled `
                    -ProvidedTo Everyone `
                    -ErrorAction Stop
                Write-Host "  Deployed Report Message add-in"
            } catch {
                Write-Host "  MarketplaceAssetId: $_"
                # Method 2: Report Phishing add-in
                try {
                    New-App -OrganizationApp -MarketplaceAssetId "WA200002469" `
                        -MarketplaceQueryMarket "en-US" `
                        -DefaultStateForUser Enabled `
                        -ProvidedTo Everyone `
                        -ErrorAction Stop
                    Write-Host "  Deployed Report Phishing add-in"
                } catch {
                    Write-Host "  Add-in deployment failed: $_"
                    Write-Host "  NOTE: Users can still report via Outlook ... > Report menu (built-in)"
                }
            }
        }
    } catch {
        Write-Host "  Add-in: $_"
    }

    # --- Step 6: Verify ---
    Write-Host ""
    Write-Host "Step 6: Verification..."

    $fp = Get-ReportSubmissionPolicy -ErrorAction SilentlyContinue
    Write-Host "  EnableReportToMicrosoft: $($fp.EnableReportToMicrosoft)"
    Write-Host "  PreSubmitMessageEnabled: $($fp.PreSubmitMessageEnabled)"
    Write-Host "  PostSubmitMessageEnabled: $($fp.PostSubmitMessageEnabled)"

    $fr = Get-ReportSubmissionRule -ErrorAction SilentlyContinue
    if ($fr) { Write-Host "  Rule: $($fr.Name) | State: $($fr.State)" }
    else { Write-Host "  Rule: Not created (built-in reporting still works)" }

    $allAddins = Get-App -OrganizationApp -ErrorAction SilentlyContinue
    $ra = $allAddins | Where-Object { $_.DisplayName -match "Report" }
    foreach ($a in $ra) { Write-Host "  Add-in: $($a.DisplayName) | Enabled: $($a.Enabled)" }

    Write-Host ""
    Write-Host "All mailboxes:"
    Get-Mailbox -ResultSize 20 | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails | Format-Table -AutoSize

    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "EXO phase complete"
} -args $user, $pass, $domain

Write-Host $exoResult

# Wait for mailbox replication
Write-Host ""
Write-Host "Waiting 30 seconds for mailbox replication..."
Start-Sleep -Seconds 30

# ================================================================
# PHASE 2: Graph API — Send 4 phishing emails from shared mailboxes
# ================================================================
Write-Host ""
Write-Host "PHASE 2: Sending phishing emails via Graph API"

$body = @{
    grant_type = "password"
    client_id  = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    scope      = "https://graph.microsoft.com/.default offline_access"
    username   = $user
    password   = $pass
}
$tok = (Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body -ContentType "application/x-www-form-urlencoded").access_token
$h = @{Authorization="Bearer $tok";"Content-Type"="application/json"}
Write-Host "Graph token acquired"

function Send-FromShared($fromAddr, $fromName, $subject, $htmlBody, $importance) {
    $imp = "normal"; if ($importance) { $imp = $importance }
    $mail = @{
        message = @{
            subject = $subject
            body = @{ contentType = "HTML"; content = $htmlBody }
            toRecipients = @(@{ emailAddress = @{ address = $user } })
            from = @{ emailAddress = @{ address = $fromAddr; name = $fromName } }
            importance = $imp
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 10
    for ($r = 0; $r -lt 3; $r++) {
        try {
            Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/me/sendMail" -Headers $h -Body $mail | Out-Null
            return $true
        } catch {
            if ($_.Exception.Response.StatusCode -eq 429) { Start-Sleep -Seconds (15*($r+1)) }
            else { Write-Host "  Error: $($_.Exception.Message)"; return $false }
        }
    }
    return $false
}

$sent = 0

# ================================================================
# TRUE POSITIVE 1: Password Expiry + Credential Harvest
# FROM: itsecurity@ | TO: odl_user
# ================================================================
Write-Host ""
Write-Host "TP-1: Password Expiry Credential Harvest..."

$tp1 = @"
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:auto;font-family:Segoe UI,Arial,sans-serif">
<tr><td style="background:#0078d4;padding:20px;text-align:center">
<img src="https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b?ver=5c31" width="120" alt="Microsoft" style="display:inline"/>
<h2 style="color:white;margin:10px 0 0">Microsoft 365 Account Security</h2>
</td></tr>
<tr><td style="padding:25px;background:#fff;border:1px solid #e0e0e0">
<p style="color:#d13438;font-size:16px;font-weight:bold">Your password expires in 2 hours</p>
<p>Dear $user,</p>
<p>Our security system detected that your Microsoft 365 password will expire at <b>11:59 PM today</b>. After expiration:</p>
<ul style="color:#333">
<li>You will be <b>locked out</b> of Outlook, Teams, and OneDrive</li>
<li>All active sessions will be <b>terminated immediately</b></li>
<li>Shared mailbox access will be <b>revoked</b></li>
<li>Your account will be <b>disabled after 24 hours</b> if not renewed</li>
</ul>
<p style="text-align:center;margin:25px 0">
<a href="http://microsoft365-password-update.secureverify-portal.com/renew?account=$user&session=aX4kT9mB2&org=$orgName&token=eyJ0eXAiOiJKV1QiLC" style="background:#0078d4;color:white;padding:14px 40px;text-decoration:none;border-radius:4px;font-size:16px;display:inline-block">Keep My Password</a>
</p>
<p style="font-size:12px;color:#666">If the button above doesn't work, copy and paste this link:

<a href="http://microsoft365-password-update.secureverify-portal.com/renew?account=$user" style="color:#0078d4;font-size:11px">https://login.microsoftonline.com/common/oauth2/authorize?client_id=00000003-0000-0ff1-ce00-0000000000&response_type=code&redirect_uri=https%3A%2F%2Fportal.office.com&scope=openid+profile</a></p>
<hr style="border:none;border-top:1px solid #eee;margin:20px 0"/>
<p style="font-size:11px;color:#999">This is an automated message from Microsoft 365 Security.

Microsoft Corporation | One Microsoft Way, Redmond, WA 98052

<a href="http://microsoft365-password-update.secureverify-portal.com/privacy" style="color:#999">Privacy Statement</a> | <a href="http://microsoft365-password-update.secureverify-portal.com/unsubscribe" style="color:#999">Unsubscribe</a>

Ref: MS365-SEC-PWD-${orgName}</p>
</td></tr></table>
"@

if (Send-FromShared "itsecurity@$domain" "Microsoft 365 Security" "Your password expires today - Immediate action required [Ticket: SEC-$(Get-Date -Format 'yyyyMMdd')]" $tp1 "high") {
    $sent++; Write-Host "  SENT"
}
Start-Sleep 4

# ================================================================
# TRUE POSITIVE 2: Fake SharePoint Document Sharing
# FROM: finance-dept@ | TO: odl_user
# ================================================================
Write-Host "TP-2: Fake SharePoint Document Sharing..."

$tp2 = @"
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:auto;font-family:Segoe UI,Arial,sans-serif">
<tr><td style="background:#0078d4;padding:15px">
<table><tr>
<td><img src="https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE1Mu3b?ver=5c31" width="28" style="vertical-align:middle"/></td>
<td style="color:white;font-size:14px;padding-left:8px"><b>SharePoint Online</b></td>
</tr></table>
</td></tr>
<tr><td style="padding:25px;background:#fff;border:1px solid #e0e0e0">
<table style="width:100%"><tr>
<td style="width:40px;vertical-align:top"><div style="background:#0078d4;color:white;width:36px;height:36px;border-radius:50%;text-align:center;line-height:36px;font-size:16px">CM</div></td>
<td style="padding-left:12px">
<p style="margin:0"><b>Catherine Mitchell</b> shared a file with you</p>
<p style="margin:2px 0;color:#666;font-size:13px">VP of Operations | Contoso Ltd</p>
</td></tr></table>

<div style="background:#f4f4f4;border:1px solid #e0e0e0;border-radius:4px;padding:15px;margin:15px 0">
<table><tr>
<td><img src="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/assets/item-types/48/xlsx.svg" width="32"/></td>
<td style="padding-left:10px">
<p style="margin:0;font-weight:bold">Q3-Salary-Review-Confidential-2026.xlsx</p>
<p style="margin:2px 0;color:#666;font-size:12px">Modified: $(Get-Date -Format 'MMMM dd, yyyy') at 9:14 AM | Size: 2.4 MB</p>
</td></tr></table>
</div>

<p><b>Message from Catherine:</b></p>
<p style="background:#fff8dc;padding:10px;border-left:3px solid #ffc107;font-style:italic">"Hi - Please review the updated salary adjustments before tomorrow's board meeting. The CEO wants final numbers by EOD. This is STRICTLY CONFIDENTIAL - do not forward. You'll need to sign in to access."</p>

<p style="text-align:center;margin:25px 0">
<a href="http://sharepoint-onedrive-$orgName.microsoftonline-files.com/documents/shared?file=Q3-Salary-Review&auth=oauth2&login_hint=$user&redirect=https%3A%2F%2F$orgName.sharepoint.com" style="background:#0078d4;color:white;padding:12px 35px;text-decoration:none;border-radius:4px;font-size:15px;display:inline-block">Open Document</a>
</p>

<p style="font-size:12px;color:#888">This link will work for <b>$user</b> only. Access expires in 24 hours.

<a href="http://sharepoint-onedrive-$orgName.microsoftonline-files.com/settings" style="color:#0078d4;font-size:11px">Manage sharing settings</a> | <a href="http://sharepoint-onedrive-$orgName.microsoftonline-files.com/block" style="color:#0078d4;font-size:11px">Block sender</a></p>
<hr style="border:none;border-top:1px solid #eee;margin:20px 0"/>
<p style="font-size:11px;color:#aaa">You're receiving this because someone shared a document with you on SharePoint.

Microsoft Corporation | One Microsoft Way, Redmond, WA 98052

<a href="http://sharepoint-onedrive-$orgName.microsoftonline-files.com/notifications" style="color:#aaa">Notification settings</a></p>
</td></tr></table>
"@

if (Send-FromShared "finance-dept@$domain" "SharePoint Online" "Catherine Mitchell shared 'Q3-Salary-Review-Confidential-2026.xlsx' with you" $tp2 "high") {
    $sent++; Write-Host "  SENT"
}
Start-Sleep 4

# ================================================================
# FALSE POSITIVE 1: Real IT Maintenance Notice
# FROM: itsecurity@ | TO: odl_user
# ================================================================
Write-Host "FP-1: Legitimate IT Maintenance..."

$fp1 = @"
<div style="max-width:600px;margin:auto;font-family:Segoe UI,Arial,sans-serif">
<p>Hi Team,</p>
<p>This is a planned maintenance notification from the IT Infrastructure team.</p>

<table style="width:100%;border-collapse:collapse;margin:15px 0">
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8;width:30%"><b>Service</b></td><td style="padding:8px;border:1px solid #ddd">Azure AD Connect Synchronization</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8"><b>Date</b></td><td style="padding:8px;border:1px solid #ddd">Saturday, $(Get-Date (Get-Date).AddDays(3) -Format 'MMMM dd, yyyy')</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8"><b>Time</b></td><td style="padding:8px;border:1px solid #ddd">2:00 AM - 6:00 AM EST</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8"><b>Impact</b></td><td style="padding:8px;border:1px solid #ddd">Password sync delays up to 4 hours</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8"><b>Action Required</b></td><td style="padding:8px;border:1px solid #ddd">None</td></tr>
</table>

<p>If you experience login issues after 6:00 AM, clear your browser cache and restart your device. For urgent issues, contact the Service Desk at ext. 4400 or Teams channel <b>#IT-Support</b>.</p>

<p>This maintenance was approved under Change Request <b>CR-2026-0847</b>.</p>

<p>Best regards,
<b>IT Infrastructure Team</b>
$orgName IT Services</p>
</div>
"@

if (Send-FromShared "itsecurity@$domain" "IT Infrastructure" "Planned Maintenance: Azure AD Connect - Saturday $(Get-Date (Get-Date).AddDays(3) -Format 'MMM d')" $fp1) {
    $sent++; Write-Host "  SENT"
}
Start-Sleep 4

# ================================================================
# FALSE POSITIVE 2: Real Team Meeting / Budget Review
# FROM: finance-dept@ | TO: odl_user
# ================================================================
Write-Host "FP-2: Legitimate Team Meeting..."

$fp2 = @"
<div style="max-width:600px;margin:auto;font-family:Segoe UI,Arial,sans-serif">
<p>Hi,</p>

<p>Quick reminder about our Q3 budget review meeting this Friday.</p>

<table style="width:100%;border-collapse:collapse;margin:15px 0">
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8;width:30%"><b>Meeting</b></td><td style="padding:8px;border:1px solid #ddd">Q3 Budget Review</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8"><b>Date</b></td><td style="padding:8px;border:1px solid #ddd">Friday, $(Get-Date (Get-Date).AddDays(2) -Format 'MMMM dd, yyyy')</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8"><b>Time</b></td><td style="padding:8px;border:1px solid #ddd">2:00 PM - 3:00 PM EST</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd;background:#f8f8f8"><b>Location</b></td><td style="padding:8px;border:1px solid #ddd">Teams (link in calendar invite)</td></tr>
</table>

<p>Updated projections based on April actuals:</p>
<ul>
<li>Marketing budget increased 12% for product launch</li>
<li>Azure consumption costs revised down 8% (reserved instances)</li>
<li>3 new Engineering headcount approved starting Q3</li>
<li>Travel budget reduced 15% per hybrid work policy update</li>
</ul>

<p>The updated spreadsheet is in our Finance SharePoint folder. The CFO will join the last 15 minutes for sign-off.</p>

<p>Please review before the meeting so we can finalize quickly.</p>

<p>Thanks,
<b>Finance Operations Team</b></p>
</div>
"@

if (Send-FromShared "finance-dept@$domain" "Finance Operations" "Reminder: Q3 Budget Review - Friday 2 PM (Updated Projections)" $fp2) {
    $sent++; Write-Host "  SENT"
}

# ================================================================
# SUMMARY
# ================================================================
Write-Host ""
Write-Host "RESULT: $sent / 4 emails sent"
Write-Host ""
Write-Host "TRUE POSITIVE (Phishing - should be reported):"
Write-Host "  TP-1: 'Your password expires today' FROM IT Security Team (itsecurity@$domain)"
Write-Host "  TP-2: 'Catherine Mitchell shared Salary-Review' FROM Finance Department (finance-dept@$domain)"
Write-Host ""
Write-Host "FALSE POSITIVE (Legitimate - should NOT be reported):"
Write-Host "  FP-1: 'Planned Maintenance: Azure AD Connect' FROM IT Security Team (itsecurity@$domain)"
Write-Host "  FP-2: 'Q3 Budget Review - Friday 2 PM' FROM Finance Department (finance-dept@$domain)"
Write-Host ""
Write-Host "In Outlook: click ... on TP emails > Report > Report phishing"
Write-Host "Check Defender portal > Email and collaboration > Submissions > User reported"

       # If successful:
        $results += [PSCustomObject]@{
            User       = $user
            Tenant     = $domain
            Status     = "Completed"
            LastRun    = Get-Date
            EmailsSent = $sent
            Error      = ""
        }

    }
    catch {

        $results += [PSCustomObject]@{
            User       = $user
            Tenant     = $domain
            Status     = "Failed"
            LastRun    = Get-Date
            EmailsSent = 0
            Error      = $_.Exception.Message
        }

        Write-Host "ERROR: $($_.Exception.Message)"
    }

}


$outputFile = ".\phishing_env9.xlsx"

$results | Export-Excel `
    -Path $outputFile `
    -WorksheetName "Cleanup Results" `
    -AutoSize `
    -BoldTopRow `
    -FreezeTopRow

