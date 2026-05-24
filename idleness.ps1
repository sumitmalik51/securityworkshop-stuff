



$serviceName = "Idle Checker Windows Service"

$idleCheckerService = Get-WmiObject `
    -Class Win32_Service `
    -Filter "Name='$serviceName'"

if ($idleCheckerService) {

    Stop-Service `
        -Name $serviceName `
        -Force `
        -ErrorAction SilentlyContinue
}

Get-Process `
    -Name "IdleCheckerUiApp" `
    -ErrorAction SilentlyContinue |
Stop-Process -Force

Get-Process `
    -Name "IdleCheckerWindowsService" `
    -ErrorAction SilentlyContinue |
Stop-Process -Force

if (Get-Service $serviceName -ErrorAction SilentlyContinue) {

    Get-CimInstance `
        -ClassName Win32_Service `
        -Filter "Name='$serviceName'" |
    Remove-CimInstance
}

$idleTrackingFolder = "C:\idle-tracking"

if (Test-Path $idleTrackingFolder) {

    Remove-Item `
        -LiteralPath $idleTrackingFolder `
        -Force `
        -Recurse
}

$idleTrackingZipFile = "C:\idle-tracking.zip"

if (Test-Path $idleTrackingZipFile) {

    Remove-Item `
        $idleTrackingZipFile `
        -Force
}

Write-Output "Idle Tracker Cleanup Completed"

