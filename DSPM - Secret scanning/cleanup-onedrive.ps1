#========================================
# Cleanup OneDrive - Remove Extra Folders
# Keeps only DSPM-Secret-Scanning-Test folder
# Deletes all other folders/files from OneDrive root
#========================================

param(
    [string]$ExcelFilePath = "C:\certs\spns.xlsx",
    [string[]]$Batches = @("Batch1", "Batch2", "Batch3", "Batch4", "Batch5"),
    [string]$LogFilePath = "C:\certs\DSPM - Secret scanning\cleanup-results.csv",
    [string]$KeepFolder = "DSPM-Secret-Scanning-Test"
)

Import-Module ImportExcel -ErrorAction Stop

function Process-User {
    param(
        [string]$User,
        [string]$Pass,
        [string]$Tenant,
        [string]$LogFile,
        [int]$UserIndex,
        [int]$TotalUsers,
        [string]$BatchName,
        [string]$Keep
    )

    Write-Host "=======================================" -ForegroundColor Green
    Write-Host "[$BatchName] USER $UserIndex / $TotalUsers" -ForegroundColor Yellow
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host "User: $User" -ForegroundColor Cyan

    # Authenticate
    try {
        $tokenBody = @{
            grant_type = "password"
            client_id  = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
            scope      = "https://graph.microsoft.com/.default offline_access"
            username   = $User
            password   = $Pass
        }
        $tokenUri = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token"
        $token = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
        $accessToken = $token.access_token
        Write-Host "  OK: Authenticated" -ForegroundColor Green
    }
    catch {
        Write-Host "  FAIL: Auth failed. Skipping." -ForegroundColor Red
        if ($LogFile) {
            [PSCustomObject]@{ Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Batch = $BatchName; User = $User; Status = "Auth Failed"; ExtraItems = 0; Deleted = 0 } | Export-Csv -Path $LogFile -Append -NoTypeInformation -Force
        }
        return
    }

    $headers = @{ "Authorization" = "Bearer $accessToken"; "Content-Type" = "application/json" }

    # List all items in OneDrive root
    try {
        $rootItems = @()
        $uri = "https://graph.microsoft.com/v1.0/me/drive/root/children?`$select=id,name,folder,file"
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
        $rootItems += $response.value

        while ($response.'@odata.nextLink') {
            $response = Invoke-RestMethod -Method Get -Uri $response.'@odata.nextLink' -Headers $headers -ErrorAction Stop
            $rootItems += $response.value
        }
    }
    catch {
        Write-Host "  FAIL: Could not list OneDrive root: $($_.Exception.Message)" -ForegroundColor Red
        if ($LogFile) {
            [PSCustomObject]@{ Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Batch = $BatchName; User = $User; Status = "List Failed"; ExtraItems = 0; Deleted = 0 } | Export-Csv -Path $LogFile -Append -NoTypeInformation -Force
        }
        return
    }

    # Find items to delete (everything except the keep folder)
    $extraItems = $rootItems | Where-Object { $_.name -ne $Keep }

    if ($extraItems.Count -eq 0) {
        Write-Host "  CLEAN: Only '$Keep' found. Nothing to delete." -ForegroundColor Green
        if ($LogFile) {
            [PSCustomObject]@{ Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Batch = $BatchName; User = $User; Status = "Clean"; ExtraItems = 0; Deleted = 0 } | Export-Csv -Path $LogFile -Append -NoTypeInformation -Force
        }
        return
    }

    Write-Host "  Found $($extraItems.Count) extra item(s) to delete:" -ForegroundColor Yellow
    $deletedCount = 0
    foreach ($item in $extraItems) {
        $itemType = if ($item.folder) { "folder" } else { "file" }
        Write-Host "    Deleting $itemType`: $($item.name)..." -ForegroundColor Yellow -NoNewline

        try {
            Invoke-RestMethod -Method Delete `
                -Uri "https://graph.microsoft.com/v1.0/me/drive/items/$($item.id)" `
                -Headers @{ "Authorization" = "Bearer $accessToken" } `
                -ErrorAction Stop
            $deletedCount++
            Write-Host " OK" -ForegroundColor Green
        }
        catch {
            Write-Host " FAIL ($($_.Exception.Message))" -ForegroundColor Red
        }
    }

    Write-Host "  Deleted $deletedCount / $($extraItems.Count) extra items" -ForegroundColor Cyan
    Write-Host ""

    if ($LogFile) {
        $status = if ($deletedCount -eq $extraItems.Count) { "Cleaned" } else { "Partial Clean" }
        [PSCustomObject]@{ Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Batch = $BatchName; User = $User; Status = $status; ExtraItems = $extraItems.Count; Deleted = $deletedCount } | Export-Csv -Path $LogFile -Append -NoTypeInformation -Force
    }
}

# ========== MAIN ==========

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "OneDrive Cleanup - Remove Extra Folders" -ForegroundColor Magenta
Write-Host "Keep: $KeepFolder" -ForegroundColor White
Write-Host "Batches: $($Batches -join ', ')" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

foreach ($batchName in $Batches) {
    Write-Host ""
    Write-Host "########################################" -ForegroundColor Magenta
    Write-Host "# PROCESSING $batchName" -ForegroundColor Magenta
    Write-Host "########################################" -ForegroundColor Magenta

    try {
        $rows = Import-Excel -Path $ExcelFilePath -WorksheetName $batchName -ErrorAction Stop
    }
    catch {
        Write-Host "WARN: Sheet '$batchName' not found. Skipping." -ForegroundColor Yellow
        continue
    }

    if (-not $rows -or $rows.Count -eq 0) { continue }

    $totalRows = $rows.Count
    Write-Host "Users in $batchName`: $totalRows" -ForegroundColor Cyan

    $idx = 0
    foreach ($r in $rows) {
        $uName = [string]$r.odluser
        $uPass = [string]$r.odlpassword
        $uTenant = [string]$r.TenantId

        if ([string]::IsNullOrWhiteSpace($uName)) { continue }

        $idx++
        Process-User -User $uName -Pass $uPass -Tenant $uTenant -LogFile $LogFilePath -UserIndex $idx -TotalUsers $totalRows -BatchName $batchName -Keep $KeepFolder
    }

    Write-Host "Finished $batchName ($idx users)" -ForegroundColor Green
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "CLEANUP COMPLETE" -ForegroundColor Yellow
Write-Host "Log: $LogFilePath" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Yellow
