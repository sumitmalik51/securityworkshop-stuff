<#
.SYNOPSIS
    Cost-reduction cleanup across ALL lab tenants from spns.xlsx.

.DESCRIPTION
    For each tenant/subscription in the specified batches:
      1. Authenticates via SPN (AppId/AppSecret)
      2. Removes resource locks on all RGs and VMs/disks
      3. Converts OS + data disks from Premium SSD -> Standard SSD
      4. Deletes Public IPs on Client VMs (VMs named ClientVM*); server PIPs preserved

    VMs are deallocated for disk SKU change, then left deallocated.

.EXAMPLE
    .\Reduce-LabCost-AllBatches.ps1 -Batches Batch1,Batch2,Batch3,Batch4,Batch5
    .\Reduce-LabCost-AllBatches.ps1 -Batches Batch1 -WhatIf

.NOTES
    Requires: Az.Accounts, Az.Compute, Az.Network, Az.Resources, ImportExcel
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]] $Batches     = @('Batch1','Batch2','Batch3','Batch4','Batch5'),
    [string]   $ExcelFile   = 'C:\certs\spns.xlsx',
    [string]   $StandardSku = 'StandardSSD_LRS',
    [switch]   $SkipPipDelete,
    [switch]   $SkipDiskConvert
)

$ErrorActionPreference = 'Continue'

# ── helpers ──────────────────────────────────────────────
function Remove-LocksOnResource {
    param([string]$ResourceGroupName, [string]$ResourceName, [string]$ResourceType)
    $locks = Get-AzResourceLock -ResourceGroupName $ResourceGroupName `
                                -ResourceName $ResourceName `
                                -ResourceType $ResourceType `
                                -ErrorAction SilentlyContinue
    foreach ($l in $locks) {
        Write-Host "    Removing lock '$($l.Name)' on $ResourceName" -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess($l.LockId, "Remove lock")) {
            Remove-AzResourceLock -LockId $l.LockId -Force | Out-Null
        }
    }
}

function Remove-RGLocks {
    param([string]$ResourceGroupName)
    $locks = Get-AzResourceLock -ResourceGroupName $ResourceGroupName -ErrorAction SilentlyContinue |
             Where-Object { -not $_.ResourceName }
    foreach ($l in $locks) {
        Write-Host "    Removing RG lock '$($l.Name)' on $ResourceGroupName" -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess($l.LockId, "Remove RG lock")) {
            Remove-AzResourceLock -LockId $l.LockId -Force | Out-Null
        }
    }
}

# ── load accounts ────────────────────────────────────────
$accounts = @()
foreach ($batch in $Batches) {
    $accounts += Import-Excel -Path $ExcelFile -WorksheetName $batch
}
$total = $accounts.Count
Write-Host "Loaded $total accounts from batches: $($Batches -join ', ')" -ForegroundColor Cyan
Write-Host ""

# ── process each tenant ──────────────────────────────────
$global:processed = 0
$global:failed    = 0
$results = @()

for ($i = 0; $i -lt $total; $i++) {
    $acct = $accounts[$i]
    $idx  = $i + 1
    $short = ($acct.odluser -split '@')[0]

    Write-Host "`n=== [$idx/$total] $short ===" -ForegroundColor Cyan

    # ── skip if no SubscriptionId ──
    if (-not $acct.SubscriptionId) {
        Write-Host "  No SubscriptionId - skipping (unprovisioned)" -ForegroundColor DarkGray
        $results += [pscustomobject]@{ Index=$idx; User=$short; Status="NO_SUB"; VMs=0; DisksConverted=0; PIPsDeleted=0 }
        $global:processed++
        continue
    }

    # ── authenticate as SPN ──
    try {
        $secSecret = ConvertTo-SecureString $acct.AppSecret -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($acct.AppId, $secSecret)
        Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $acct.TenantId -ErrorAction Stop | Out-Null
        Select-AzSubscription -SubscriptionId $acct.SubscriptionId -ErrorAction Stop | Out-Null
        Write-Host "  Authenticated to subscription $($acct.SubscriptionId)" -ForegroundColor Green
    } catch {
        Write-Host "  FAILED to authenticate: $_" -ForegroundColor Red
        $global:failed++
        $results += [pscustomobject]@{ Index=$idx; User=$short; Status="AUTH_FAILED"; VMs=0; DisksConverted=0; PIPsDeleted=0 }
        continue
    }

    # ── discover all RGs ──
    $rgs = Get-AzResourceGroup -ErrorAction SilentlyContinue
    if (-not $rgs) {
        Write-Host "  No resource groups found - skipping" -ForegroundColor DarkGray
        $results += [pscustomobject]@{ Index=$idx; User=$short; Status="NO_RGs"; VMs=0; DisksConverted=0; PIPsDeleted=0 }
        $global:processed++
        continue
    }

    # ── Step 1: Remove locks on all RGs ──
    Write-Host "  Step 1: Removing resource locks..." -ForegroundColor White
    foreach ($rg in $rgs) {
        Remove-RGLocks -ResourceGroupName $rg.ResourceGroupName
    }

    # ── discover all VMs ──
    $allVMs = Get-AzVM -ErrorAction SilentlyContinue
    if (-not $allVMs -or $allVMs.Count -eq 0) {
        Write-Host "  No VMs found - skipping disk/PIP steps" -ForegroundColor DarkGray
        $results += [pscustomobject]@{ Index=$idx; User=$short; Status="NO_VMs"; VMs=0; DisksConverted=0; PIPsDeleted=0 }
        $global:processed++
        continue
    }
    Write-Host "  Found $($allVMs.Count) VM(s): $($allVMs.Name -join ', ')" -ForegroundColor Gray

    $disksConverted = 0
    $pipsDeleted    = 0

    foreach ($vm in $allVMs) {
        $vmRG = $vm.ResourceGroupName

        # Remove locks on VM
        Remove-LocksOnResource -ResourceGroupName $vmRG -ResourceName $vm.Name -ResourceType 'Microsoft.Compute/virtualMachines'

        # Remove locks on OS disk
        $osDiskId   = $vm.StorageProfile.OsDisk.ManagedDisk.Id
        $osDiskName = ($osDiskId -split '/')[-1]
        $diskRG     = ($osDiskId -split '/')[4]
        Remove-LocksOnResource -ResourceGroupName $diskRG -ResourceName $osDiskName -ResourceType 'Microsoft.Compute/disks'

        # ── Step 2: Convert disks ──
        if (-not $SkipDiskConvert) {
            # Check VM power state - must be deallocated for SKU change
            $vmStatus = (Get-AzVM -ResourceGroupName $vmRG -Name $vm.Name -Status -ErrorAction SilentlyContinue).Statuses |
                        Where-Object { $_.Code -like 'PowerState/*' } |
                        Select-Object -ExpandProperty Code
            $needDealloc = $false

            # OS disk
            $osDisk = Get-AzDisk -ResourceGroupName $diskRG -DiskName $osDiskName -ErrorAction SilentlyContinue
            if ($osDisk -and $osDisk.Sku.Name -ne $StandardSku) {
                # Deallocate if needed (once per VM)
                if ($vmStatus -ne 'PowerState/deallocated' -and -not $needDealloc) {
                    Write-Host "    $($vm.Name) state: $vmStatus -> deallocating" -ForegroundColor Yellow
                    if ($PSCmdlet.ShouldProcess($vm.Name, "Stop-AzVM -Force")) {
                        Stop-AzVM -ResourceGroupName $vmRG -Name $vm.Name -Force -ErrorAction SilentlyContinue | Out-Null
                    }
                    $needDealloc = $true
                }
                Write-Host "    $osDiskName : $($osDisk.Sku.Name) -> $StandardSku" -ForegroundColor Gray
                $osDisk.Sku = New-Object Microsoft.Azure.Management.Compute.Models.DiskSku($StandardSku)
                if ($PSCmdlet.ShouldProcess($osDiskName, "Convert to $StandardSku")) {
                    Update-AzDisk -ResourceGroupName $diskRG -DiskName $osDiskName -Disk $osDisk -ErrorAction SilentlyContinue | Out-Null
                    $disksConverted++
                }
            } else {
                Write-Host "    $osDiskName : already $StandardSku" -ForegroundColor DarkGray
            }

            # Data disks
            foreach ($dd in $vm.StorageProfile.DataDisks) {
                $ddId   = $dd.ManagedDisk.Id
                $ddName = ($ddId -split '/')[-1]
                $ddRG   = ($ddId -split '/')[4]
                $ddObj  = Get-AzDisk -ResourceGroupName $ddRG -DiskName $ddName -ErrorAction SilentlyContinue
                if ($ddObj -and $ddObj.Sku.Name -ne $StandardSku) {
                    if ($vmStatus -ne 'PowerState/deallocated' -and -not $needDealloc) {
                        if ($PSCmdlet.ShouldProcess($vm.Name, "Stop-AzVM -Force")) {
                            Stop-AzVM -ResourceGroupName $vmRG -Name $vm.Name -Force -ErrorAction SilentlyContinue | Out-Null
                        }
                        $needDealloc = $true
                    }
                    Write-Host "    data disk $ddName : $($ddObj.Sku.Name) -> $StandardSku" -ForegroundColor Gray
                    $ddObj.Sku = New-Object Microsoft.Azure.Management.Compute.Models.DiskSku($StandardSku)
                    if ($PSCmdlet.ShouldProcess($ddName, "Convert data disk to $StandardSku")) {
                        Update-AzDisk -ResourceGroupName $ddRG -DiskName $ddName -Disk $ddObj -ErrorAction SilentlyContinue | Out-Null
                        $disksConverted++
                    }
                }
            }
        }

        # ── Step 3: Delete PIPs on Client VMs only ──
        if (-not $SkipPipDelete -and $vm.Name -like 'ClientVM*') {
            foreach ($nicRef in $vm.NetworkProfile.NetworkInterfaces) {
                $nicName = ($nicRef.Id -split '/')[-1]
                $nicRG   = ($nicRef.Id -split '/')[4]
                $nic     = Get-AzNetworkInterface -ResourceGroupName $nicRG -Name $nicName -ErrorAction SilentlyContinue
                if (-not $nic) { continue }

                $pipsToDelete = @()
                $modified = $false

                foreach ($ipc in $nic.IpConfigurations) {
                    if ($ipc.PublicIpAddress) {
                        $pipName = ($ipc.PublicIpAddress.Id -split '/')[-1]
                        $pipRG   = ($ipc.PublicIpAddress.Id -split '/')[4]
                        Remove-LocksOnResource -ResourceGroupName $pipRG -ResourceName $pipName -ResourceType 'Microsoft.Network/publicIPAddresses'
                        Write-Host "    Dissociating PIP '$pipName' from NIC '$nicName'" -ForegroundColor Yellow
                        $ipc.PublicIpAddress = $null
                        $modified = $true
                        $pipsToDelete += @{ Name = $pipName; RG = $pipRG }
                    }
                }

                if ($modified -and $PSCmdlet.ShouldProcess($nicName, "Update NIC (dissociate PIPs)")) {
                    Set-AzNetworkInterface -NetworkInterface $nic -ErrorAction SilentlyContinue | Out-Null
                }

                foreach ($p in $pipsToDelete) {
                    if ($PSCmdlet.ShouldProcess($p.Name, "Delete PIP")) {
                        Remove-AzPublicIpAddress -ResourceGroupName $p.RG -Name $p.Name -Force -ErrorAction SilentlyContinue | Out-Null
                        Write-Host "    Deleted PIP '$($p.Name)'" -ForegroundColor Green
                        $pipsDeleted++
                    }
                }
            }
        }
    }

    $results += [pscustomobject]@{ Index=$idx; User=$short; Status="OK"; VMs=$allVMs.Count; DisksConverted=$disksConverted; PIPsDeleted=$pipsDeleted }
    $global:processed++
    Write-Host "  DONE - $($allVMs.Count) VMs, $disksConverted disks converted, $pipsDeleted PIPs deleted" -ForegroundColor Green
}

# ── summary ──────────────────────────────────────────────
Write-Host "`n=======================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "Total:     $total"
Write-Host "Processed: $($global:processed)" -ForegroundColor Green
$failColor = if ($global:failed -gt 0) { 'Red' } else { 'Green' }
Write-Host "Failed:    $($global:failed)" -ForegroundColor $failColor
$results | Format-Table -AutoSize
$batchLabel = ($Batches -join '-')
$csvPath = "C:\certs\LabCost-Reduction-Results-$batchLabel.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "Results saved to $csvPath" -ForegroundColor Cyan
