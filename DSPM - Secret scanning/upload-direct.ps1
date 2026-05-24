#========================================
# Upload Test Secrets to User OneDrive
# Direct Graph API Upload with User Credentials
# Checks if folder already uploaded; skips if complete
# Processes all batches from spns.xlsx
#========================================

param(
    [string]$Username = "",
    [string]$Password = "",
    [string]$TenantId = "",
    [string]$ExcelFilePath = "C:\certs\spns.xlsx",
    [string[]]$Batches = @("Batch1", "Batch2", "Batch3", "Batch4", "Batch5"),
    [string]$LogFilePath = "C:\certs\DSPM - Secret scanning\upload-results.csv"
)

$sourceFolder = "C:\certs\DSPM - Secret scanning\DSPM-Secret-Scanning-Test"
$expectedFileCount = (Get-ChildItem $sourceFolder -File).Count

# Helper function to process a single user
function Process-User {
    param(
        [string]$User,
        [string]$Pass,
        [string]$Tenant,
        [string]$LogFile,
        [int]$UserIndex,
        [int]$TotalUsers,
        [string]$BatchName
    )
    
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host "[$BatchName] USER $UserIndex / $TotalUsers" -ForegroundColor Yellow
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "User: $User" -ForegroundColor Cyan
    Write-Host "Authenticating..." -ForegroundColor Green

    # Step 1: Get access token
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
        Write-Host "OK: Authentication successful" -ForegroundColor Green
    }
    catch {
        Write-Host "FAIL: Authentication failed for $User. Skipping..." -ForegroundColor Red
        if ($LogFile) {
            [PSCustomObject]@{ Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Batch = $BatchName; User = $User; Status = "Auth Failed"; CloudFiles = 0; Expected = $script:expectedFileCount; Uploaded = 0; Failed = 0 } | Export-Csv -Path $LogFile -Append -NoTypeInformation -Force
        }
        return
    }

    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Content-Type"  = "application/json"
    }

    $folderName = "DSPM-Secret-Scanning-Test"

    # Step 2: Check if folder already exists and has all files
    Write-Host ""
    Write-Host "[1] Checking if folder already uploaded..." -ForegroundColor Green

    $folderId = $null
    $existingFileNames = @()

    try {
        $folderCheck = Invoke-RestMethod -Method Get `
            -Uri "https://graph.microsoft.com/v1.0/me/drive/root:/${folderName}" `
            -Headers $headers `
            -ErrorAction Stop
        
        $folderId = $folderCheck.id
        Write-Host "  Folder exists. Checking files..." -ForegroundColor Yellow

        # List all files in the folder
        $childrenUri = "https://graph.microsoft.com/v1.0/me/drive/items/$folderId/children?`$top=200&`$select=name"
        $childrenResponse = Invoke-RestMethod -Method Get -Uri $childrenUri -Headers $headers -ErrorAction Stop
        $existingFileNames = @($childrenResponse.value | ForEach-Object { $_.name })

        # Handle pagination
        while ($childrenResponse.'@odata.nextLink') {
            $childrenResponse = Invoke-RestMethod -Method Get -Uri $childrenResponse.'@odata.nextLink' -Headers $headers -ErrorAction Stop
            $existingFileNames += @($childrenResponse.value | ForEach-Object { $_.name })
        }

        $cloudCount = $existingFileNames.Count
        Write-Host "  Cloud files: $cloudCount / $($script:expectedFileCount) expected" -ForegroundColor Cyan

        if ($cloudCount -ge $script:expectedFileCount) {
            Write-Host "  SKIP: All $cloudCount files already uploaded for $User" -ForegroundColor Green
            Write-Host ""
            if ($LogFile) {
                [PSCustomObject]@{ Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Batch = $BatchName; User = $User; Status = "Skipped - Already Complete"; CloudFiles = $cloudCount; Expected = $script:expectedFileCount; Uploaded = 0; Failed = 0 } | Export-Csv -Path $LogFile -Append -NoTypeInformation -Force
            }
            return
        }

        Write-Host "  Missing files detected. Will upload missing ones..." -ForegroundColor Yellow
    }
    catch {
        Write-Host "  Folder not found. Will create and upload all files." -ForegroundColor Yellow
    }

    # Step 3: Create folder if it doesn't exist
    if (-not $folderId) {
        Write-Host ""
        Write-Host "[2] Creating folder in OneDrive..." -ForegroundColor Green

        $createFolderBody = @{
            name                                = $folderName
            folder                              = @{}
            "@microsoft.graph.conflictBehavior" = "replace"
        } | ConvertTo-Json

        try {
            $folderResponse = Invoke-RestMethod -Method Post `
                -Uri "https://graph.microsoft.com/v1.0/me/drive/root/children" `
                -Headers $headers `
                -Body $createFolderBody `
                -ErrorAction Stop
            
            $folderId = $folderResponse.id
            Write-Host "  OK: Folder created: $folderName" -ForegroundColor Green
        }
        catch {
            Write-Host "  FAIL: Could not create folder: $_" -ForegroundColor Red
            if ($LogFile) {
                [PSCustomObject]@{ Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Batch = $BatchName; User = $User; Status = "Folder Create Failed"; CloudFiles = 0; Expected = $script:expectedFileCount; Uploaded = 0; Failed = 0 } | Export-Csv -Path $LogFile -Append -NoTypeInformation -Force
            }
            return
        }
    }

    # Step 4: Upload only missing files
    Write-Host ""
    Write-Host "[3] Uploading missing files to cloud..." -ForegroundColor Green

    $allFiles = Get-ChildItem $sourceFolder -File
    # Filter to only files not already in cloud
    $filesToUpload = $allFiles | Where-Object { $_.Name -notin $existingFileNames }

    Write-Host "  Total source files: $($allFiles.Count), Already in cloud: $($existingFileNames.Count), To upload: $($filesToUpload.Count)" -ForegroundColor Cyan

    $uploadedCount = 0
    $failedCount = 0
    $failedFiles = @()
    $totalToUpload = $filesToUpload.Count

    foreach ($file in $filesToUpload) {
        $uploadedCount++
        $fileName = $file.Name
        
        try {
            $uploadUri = "https://graph.microsoft.com/v1.0/me/drive/items/$folderId`:/$fileName`:/content"
            $uploadHeaders = @{ "Authorization" = "Bearer $accessToken" }

            $null = Invoke-RestMethod -Method Put `
                -Uri $uploadUri `
                -Headers $uploadHeaders `
                -InFile $file.FullName `
                -ContentType "application/octet-stream" `
                -ErrorAction Stop

            Write-Host "  [$uploadedCount/$totalToUpload] OK: $fileName" -ForegroundColor Green
        }
        catch {
            $failedCount++
            $failedFiles += $fileName
            $statusCode = $null
            try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch {}
            if ($statusCode) {
                Write-Host "  [$uploadedCount/$totalToUpload] FAIL: $fileName (HTTP $statusCode)" -ForegroundColor Red
            }
            else {
                Write-Host "  [$uploadedCount/$totalToUpload] FAIL: $fileName ($($_.Exception.Message))" -ForegroundColor Red
            }
        }
    }

    Write-Host ""
    Write-Host "=======================================" -ForegroundColor Green
    if ($failedCount -eq 0) {
        Write-Host "OK: UPLOAD COMPLETE FOR $User!" -ForegroundColor Yellow
    }
    else {
        Write-Host "WARNING: UPLOAD COMPLETED WITH ERRORS FOR $User" -ForegroundColor Yellow
    }
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "  Account: $User" -ForegroundColor White
    Write-Host "  Already in cloud: $($existingFileNames.Count)" -ForegroundColor White
    Write-Host "  Newly uploaded: $($uploadedCount - $failedCount)" -ForegroundColor White
    Write-Host "  Failed: $failedCount" -ForegroundColor White

    if ($failedCount -gt 0) {
        Write-Host "  First failed files:" -ForegroundColor Yellow
        $failedFiles | Select-Object -First 10 | ForEach-Object { Write-Host "    - $_" -ForegroundColor White }
    }
    Write-Host ""

    if ($LogFile) {
        $status = if ($failedCount -eq 0 -and $uploadedCount -gt 0) { "Uploaded" } elseif ($uploadedCount -eq 0) { "Nothing To Upload" } else { "Partial Upload" }
        [PSCustomObject]@{ Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Batch = $BatchName; User = $User; Status = $status; CloudFiles = ($existingFileNames.Count + $uploadedCount - $failedCount); Expected = $script:expectedFileCount; Uploaded = ($uploadedCount - $failedCount); Failed = $failedCount } | Export-Csv -Path $LogFile -Append -NoTypeInformation -Force
    }
}

# ========== MAIN ==========

if ($Username -and $Password -and $TenantId) {
    Process-User -User $Username -Pass $Password -Tenant $TenantId -LogFile $LogFilePath -UserIndex 1 -TotalUsers 1 -BatchName "Single"
}
elseif (Test-Path $ExcelFilePath) {
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "DSPM Secret Scanning - Upload to OneDrive" -ForegroundColor Magenta
    Write-Host "Source: $sourceFolder" -ForegroundColor White
    Write-Host "Expected files per user: $expectedFileCount" -ForegroundColor White
    Write-Host "Batches: $($Batches -join ', ')" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""

    Import-Module ImportExcel -ErrorAction Stop

    foreach ($batchName in $Batches) {
        Write-Host ""
        Write-Host "########################################" -ForegroundColor Magenta
        Write-Host "# PROCESSING $batchName" -ForegroundColor Magenta
        Write-Host "########################################" -ForegroundColor Magenta

        try {
            $rows = Import-Excel -Path $ExcelFilePath -WorksheetName $batchName -ErrorAction Stop
        }
        catch {
            Write-Host "WARN: Sheet '$batchName' not found or error reading. Skipping." -ForegroundColor Yellow
            continue
        }

        if (-not $rows -or $rows.Count -eq 0) {
            Write-Host "WARN: No data in $batchName. Skipping." -ForegroundColor Yellow
            continue
        }

        $totalRows = $rows.Count
        Write-Host "Users in $batchName`: $totalRows" -ForegroundColor Cyan
        Write-Host ""

        $idx = 0
        foreach ($r in $rows) {
            $uName = [string]$r.odluser
            $uPass = [string]$r.odlpassword
            $uTenant = [string]$r.TenantId

            if ([string]::IsNullOrWhiteSpace($uName)) { continue }

            $idx++
            Process-User -User $uName -Pass $uPass -Tenant $uTenant -LogFile $LogFilePath -UserIndex $idx -TotalUsers $totalRows -BatchName $batchName
        }

        Write-Host ""
        Write-Host "Finished $batchName ($idx users processed)" -ForegroundColor Green
    }
}
else {
    Write-Host "Please provide either Username/Password/TenantId or a valid ExcelFilePath." -ForegroundColor Red
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ALL BATCHES COMPLETE" -ForegroundColor Yellow
Write-Host "Log file: $LogFilePath" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "WAIT: Wait 15-20 minutes for indexing, then:" -ForegroundColor Yellow
Write-Host "  1. Go to Purview: https://purview.microsoft.com" -ForegroundColor White
Write-Host "  2. DSPM > Discover > Asset explorer" -ForegroundColor White
Write-Host "  3. Select: Posture Agent (preview)" -ForegroundColor White
Write-Host "  4. Run credential scanning" -ForegroundColor White
