$env:AZURE_ENABLE_WAM='false'
[System.Environment]::SetEnvironmentVariable('Broker_Enabled','false','Process')
$sp = ConvertTo-SecureString 'zpcj65FDY*yM' -AsPlainText -Force
$cr = New-Object System.Management.Automation.PSCredential('odl_user_2226966@otuwacne103898.onmicrosoft.com', $sp)
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -Credential $cr -ShowBanner:$false -ErrorAction Stop
Write-Host 'Connected'
$c = Get-AdminAuditLogConfig
Write-Host "Current: $($c.UnifiedAuditLogIngestionEnabled)"
try {
    Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -ErrorAction Stop
    Write-Host 'SET OK'
} catch {
    Write-Host "SET FAIL: $($_.Exception.Message)"
}
Start-Sleep 5
$c2 = Get-AdminAuditLogConfig
Write-Host "After: $($c2.UnifiedAuditLogIngestionEnabled)"
Disconnect-ExchangeOnline -Confirm:$false -EA SilentlyContinue
