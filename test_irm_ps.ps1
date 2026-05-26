# Test IRM Policy check via Security & Compliance PowerShell
# Uses Connect-IPPSSession to connect to Security & Compliance Center
# Then runs Get-InsiderRiskPolicy

param(
    [string]$Sheet = "Batch1",
    [int]$RowIndex = 0
)

Import-Module ExchangeOnlineManagement -MinimumVersion 3.0

# Load credentials from spns.xlsx
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open("C:\certs\spns.xlsx")
$worksheet = $workbook.Sheets.Item($Sheet)

# Find column indices (row 1 is header)
$headers = @{}
$col = 1
while ($worksheet.Cells.Item(1, $col).Value2) {
    $headers[$worksheet.Cells.Item(1, $col).Value2] = $col
    $col++
}

$dataRow = $RowIndex + 2  # +2 for 1-based + header row
$username = $worksheet.Cells.Item($dataRow, $headers["odluser"]).Value2
$password = $worksheet.Cells.Item($dataRow, $headers["odlpassword"]).Value2
$tenantId = $worksheet.Cells.Item($dataRow, $headers["TenantId"]).Value2

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "User: $username"
Write-Host "TenantID: $tenantId"

# Create credential
$secPwd = ConvertTo-SecureString $password -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($username, $secPwd)

# Connect to Security & Compliance Center
Write-Host "`nConnecting to Security & Compliance Center..."
try {
    Connect-IPPSSession -Credential $cred -ErrorAction Stop
    Write-Host "Connected successfully!"
    
    # Get IRM Policies
    Write-Host "`nGetting Insider Risk Management Policies..."
    $policies = Get-InsiderRiskPolicy -ErrorAction Stop
    
    if ($policies) {
        Write-Host "`nFound $($policies.Count) IRM policies:"
        $policies | ForEach-Object {
            Write-Host "  Name: $($_.Name)"
            Write-Host "  DisplayName: $($_.DisplayName)"
            Write-Host "  Enabled: $($_.IsEnabled)"
            Write-Host "  ---"
        }
        # Full details
        $policies | Format-List *
    } else {
        Write-Host "No IRM policies found."
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}
