$env:AZURE_ENABLE_WAM='false'
[System.Environment]::SetEnvironmentVariable('Broker_Enabled','false','Process')
Import-Module ExchangeOnlineManagement
$sp = ConvertTo-SecureString 'jpqz62ZBO*MH' -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential('odl_user_2226954@otuwacne103891.onmicrosoft.com', $sp)
Connect-ExchangeOnline -Credential $cred -ShowBanner:$false
Write-Host 'Connected to Exchange Online'

Write-Host 'Step 1: Creating shared mailboxes...'
try { New-Mailbox -Shared -Name 'IT Security Team' -DisplayName 'IT Security Team' -Alias itsecurity -PrimarySmtpAddress itsecurity@otuwacne103891.onmicrosoft.com -ErrorAction Stop; Write-Host '  Created itsecurity@otuwacne103891.onmicrosoft.com' } catch { if ($_.Exception.Message -match 'already in use') { Write-Host '  itsecurity@otuwacne103891.onmicrosoft.com already exists' } else { Write-Host "  Error: $($_.Exception.Message)" } }
try { New-Mailbox -Shared -Name 'Finance Department' -DisplayName 'Finance Department' -Alias 'finance-dept' -PrimarySmtpAddress finance-dept@otuwacne103891.onmicrosoft.com -ErrorAction Stop; Write-Host '  Created finance-dept@otuwacne103891.onmicrosoft.com' } catch { if ($_.Exception.Message -match 'already in use') { Write-Host '  finance-dept@otuwacne103891.onmicrosoft.com already exists' } else { Write-Host "  Error: $($_.Exception.Message)" } }

Write-Host 'Step 2: Granting FullAccess + SendAs permissions...'
$mbs = @("itsecurity@otuwacne103891.onmicrosoft.com","finance-dept@otuwacne103891.onmicrosoft.com")
foreach ($mb in $mbs) {
    try { Add-MailboxPermission -Identity $mb -User 'odl_user_2226954@otuwacne103891.onmicrosoft.com' -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop | Out-Null; Write-Host "  FullAccess on $mb" } catch { Write-Host "  FullAccess: $($_.Exception.Message)" }
    try { Add-RecipientPermission -Identity $mb -Trustee 'odl_user_2226954@otuwacne103891.onmicrosoft.com' -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null; Write-Host "  SendAs on $mb" } catch { Write-Host "  SendAs: $($_.Exception.Message)" }
}

Write-Host 'Step 3: Configuring Report Submission Policy...'
try { $pol = Get-ReportSubmissionPolicy -ErrorAction SilentlyContinue; if (-not $pol) { New-ReportSubmissionPolicy -EnableReportToMicrosoft $true -EnableThirdPartyAddress $false -ReportJunkToCustomizedAddress $false -ReportNotJunkToCustomizedAddress $false -ReportPhishToCustomizedAddress $false -PreSubmitMessageEnabled $true -PostSubmitMessageEnabled $true -ErrorAction Stop; Write-Host '  Created Policy' } else { Write-Host '  Policy already exists' } } catch { Write-Host "  Error: $($_.Exception.Message)" }

Write-Host 'Step 4: Creating Report Submission Rule...'
try { $rule = Get-ReportSubmissionRule -ErrorAction SilentlyContinue; if (-not $rule) { New-ReportSubmissionRule -Name DefaultReportSubmissionRule -ReportSubmissionPolicy DefaultReportSubmissionPolicy -SentTo 'odl_user_2226954@otuwacne103891.onmicrosoft.com' -ErrorAction Stop; Write-Host '  Rule created' } else { Write-Host '  Rule already exists' } } catch { Write-Host "  Error: $($_.Exception.Message)" }

Write-Host 'Step 5: Verification...'
$p2 = Get-ReportSubmissionPolicy
Write-Host "  EnableReportToMicrosoft: $($p2.EnableReportToMicrosoft)"
$r2 = Get-ReportSubmissionRule -ErrorAction SilentlyContinue
if ($r2) { Write-Host "  Rule: $($r2.Name) | State: $($r2.State)" } else { Write-Host '  Rule: Not created' }
Write-Host 'EXO phase complete'
Disconnect-ExchangeOnline -Confirm:$false
