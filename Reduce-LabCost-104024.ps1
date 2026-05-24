<#
.SYNOPSIS
    Cost-reduction cleanup for tenant OTU WA CNE - 104024 lab resources.

.DESCRIPTION
    1. Removes resource locks on target VMs / disks / RGs
    2. Converts OS (and any data) disks from Premium SSD -> Standard SSD in-place
    3. Deletes Public IPs of CLIENT VMs only (server VM PIP is preserved)

    VMs must be deallocated for SKU change (script will deallocate if needed).

.NOTES
    Requires: Az.Accounts, Az.Compute, Az.Network, Az.Resources
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]   $SubscriptionId        = "",
    [string[]] $ClientVMs             = @('ClientVM1','ClientVM2'),
    [string]   $ClientVMResourceGroup = 'Intune-Lab-VMs',
    [string]   $ServerVM              = 'svrvm-2227104',
    [string]   $ServerVMResourceGroup = 'microsoft',
    [string]   $StandardSku           = 'StandardSSD_LRS',
    [switch]   $SkipPipDelete
)

$ErrorActionPreference = 'Stop'

function Write-Section($t) { Write-Host "`n========== $t ==========" -ForegroundColor Cyan }

function Remove-LocksOn {
    param([string]$ResourceGroup, [string]$ResourceName, [string]$ResourceType)
    $locks = Get-AzResourceLock -ResourceGroupName $ResourceGroup `
                                -ResourceName $ResourceName `
                                -ResourceType $ResourceType `
                                -ErrorAction SilentlyContinue
    foreach ($l in $locks) {
        Write-Host "  Removing lock '$($l.Name)' on $ResourceName" -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess($l.LockId, "Remove lock")) {
            Remove-AzResourceLock -LockId $l.LockId -Force | Out-Null
        }
    }
}

function Remove-RGLocks {
    param([string]$ResourceGroup)
    $locks = Get-AzResourceLock -ResourceGroupName $ResourceGroup -ErrorAction SilentlyContinue |
             Where-Object { -not $_.ResourceName }
    foreach ($l in $locks) {
        Write-Host "  Removing RG-scope lock '$($l.Name)' on $ResourceGroup" -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess($l.LockId, "Remove RG lock")) {
            Remove-AzResourceLock -LockId $l.LockId -Force | Out-Null
        }
    }
}

# ---------- login ----------
if (-not (Get-AzContext)) { Connect-AzAccount | Out-Null }
if ($SubscriptionId) { Select-AzSubscription -SubscriptionId $SubscriptionId | Out-Null }
$ctx = Get-AzContext
Write-Host "Subscription: $($ctx.Subscription.Name) ($($ctx.Subscription.Id))" -ForegroundColor Green

# ---------- targets ----------
$targets = @()
foreach ($n in $ClientVMs) {
    $targets += [pscustomobject]@{ VMName = $n; RG = $ClientVMResourceGroup; IsClient = $true }
}
$targets += [pscustomobject]@{ VMName = $ServerVM; RG = $ServerVMResourceGroup; IsClient = $false }

# ---------- 1. locks ----------
Write-Section "Step 1: Remove resource locks"
foreach ($rg in ($targets.RG | Select-Object -Unique)) { Remove-RGLocks -ResourceGroup $rg }

foreach ($t in $targets) {
    $vm = Get-AzVM -ResourceGroupName $t.RG -Name $t.VMName -ErrorAction SilentlyContinue
    if (-not $vm) { Write-Host "  VM '$($t.VMName)' not found in RG '$($t.RG)' - skipping" -ForegroundColor Red; continue }

    Remove-LocksOn -ResourceGroup $t.RG -ResourceName $t.VMName -ResourceType 'Microsoft.Compute/virtualMachines'

    $osDiskId   = $vm.StorageProfile.OsDisk.ManagedDisk.Id
    $osDiskName = ($osDiskId -split '/')[-1]
    $diskRG     = ($osDiskId -split '/')[4]
    Remove-LocksOn -ResourceGroup $diskRG -ResourceName $osDiskName -ResourceType 'Microsoft.Compute/disks'
}

# ---------- 2. convert disks ----------
Write-Section "Step 2: Convert disks -> $StandardSku"
$summary = @()
foreach ($t in $targets) {
    $vm = Get-AzVM -ResourceGroupName $t.RG -Name $t.VMName -ErrorAction SilentlyContinue
    if (-not $vm) { continue }

    $status = (Get-AzVM -ResourceGroupName $t.RG -Name $t.VMName -Status).Statuses |
              Where-Object { $_.Code -like 'PowerState/*' } | Select-Object -ExpandProperty Code
    if ($status -ne 'PowerState/deallocated') {
        Write-Host "  $($t.VMName) state: $status -> deallocating" -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess($t.VMName, "Stop-AzVM -Force")) {
            Stop-AzVM -ResourceGroupName $t.RG -Name $t.VMName -Force | Out-Null
        }
    }

    # OS disk
    $osDiskId   = $vm.StorageProfile.OsDisk.ManagedDisk.Id
    $osDiskName = ($osDiskId -split '/')[-1]
    $diskRG     = ($osDiskId -split '/')[4]
    $disk       = Get-AzDisk -ResourceGroupName $diskRG -DiskName $osDiskName
    $before     = $disk.Sku.Name
    Write-Host "  $osDiskName : $before ($($disk.DiskSizeGB) GB)" -ForegroundColor Gray
    if ($before -ne $StandardSku) {
        $disk.Sku = New-Object Microsoft.Azure.Management.Compute.Models.DiskSku($StandardSku)
        if ($PSCmdlet.ShouldProcess($osDiskName, "Update SKU -> $StandardSku")) {
            Update-AzDisk -ResourceGroupName $diskRG -DiskName $osDiskName -Disk $disk | Out-Null
            Write-Host "    -> converted to $StandardSku" -ForegroundColor Green
        }
    } else {
        Write-Host "    -> already $StandardSku" -ForegroundColor DarkGray
    }
    $summary += [pscustomobject]@{ VM=$t.VMName; Disk=$osDiskName; SizeGB=$disk.DiskSizeGB; Before=$before; After=$StandardSku; IsClient=$t.IsClient }

    # Data disks
    foreach ($dd in $vm.StorageProfile.DataDisks) {
        $ddId   = $dd.ManagedDisk.Id
        $ddName = ($ddId -split '/')[-1]
        $ddRG   = ($ddId -split '/')[4]
        $ddObj  = Get-AzDisk -ResourceGroupName $ddRG -DiskName $ddName
        $ddBefore = $ddObj.Sku.Name
        if ($ddBefore -ne $StandardSku) {
            $ddObj.Sku = New-Object Microsoft.Azure.Management.Compute.Models.DiskSku($StandardSku)
            if ($PSCmdlet.ShouldProcess($ddName, "Update data disk SKU -> $StandardSku")) {
                Update-AzDisk -ResourceGroupName $ddRG -DiskName $ddName -Disk $ddObj | Out-Null
                Write-Host "  data disk $ddName -> $StandardSku" -ForegroundColor Green
            }
        }
        $summary += [pscustomobject]@{ VM=$t.VMName; Disk=$ddName; SizeGB=$ddObj.DiskSizeGB; Before=$ddBefore; After=$StandardSku; IsClient=$t.IsClient }
    }
}

# ---------- 3. delete client PIPs ----------
if (-not $SkipPipDelete) {
    Write-Section "Step 3: Delete CLIENT VM Public IPs"
    foreach ($t in $targets | Where-Object { $_.IsClient }) {
        $vm = Get-AzVM -ResourceGroupName $t.RG -Name $t.VMName -ErrorAction SilentlyContinue
        if (-not $vm) { continue }

        foreach ($nicRef in $vm.NetworkProfile.NetworkInterfaces) {
            $nicId   = $nicRef.Id
            $nicName = ($nicId -split '/')[-1]
            $nicRG   = ($nicId -split '/')[4]
            $nic     = Get-AzNetworkInterface -ResourceGroupName $nicRG -Name $nicName
            $pipsToDelete = @()
            $modified = $false

            foreach ($ipc in $nic.IpConfigurations) {
                if ($ipc.PublicIpAddress) {
                    $pipId   = $ipc.PublicIpAddress.Id
                    $pipName = ($pipId -split '/')[-1]
                    $pipRG   = ($pipId -split '/')[4]
                    Remove-LocksOn -ResourceGroup $pipRG -ResourceName $pipName -ResourceType 'Microsoft.Network/publicIPAddresses'
                    Write-Host "  Dissociating PIP '$pipName' from NIC '$nicName'" -ForegroundColor Yellow
                    $ipc.PublicIpAddress = $null
                    $modified = $true
                    $pipsToDelete += @{ Name = $pipName; RG = $pipRG }
                }
            }

            if ($modified -and $PSCmdlet.ShouldProcess($nicName, "Update NIC (dissociate PIPs)")) {
                Set-AzNetworkInterface -NetworkInterface $nic | Out-Null
            }

            foreach ($p in $pipsToDelete) {
                Write-Host "  Deleting PIP '$($p.Name)'" -ForegroundColor Yellow
                if ($PSCmdlet.ShouldProcess($p.Name, "Remove-AzPublicIpAddress")) {
                    Remove-AzPublicIpAddress -ResourceGroupName $p.RG -Name $p.Name -Force | Out-Null
                    Write-Host "    -> deleted" -ForegroundColor Green
                }
            }
        }
    }
}

Write-Section "Summary"
$summary | Format-Table -AutoSize
Write-Host "Done." -ForegroundColor Green
