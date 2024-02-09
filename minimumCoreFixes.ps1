<#
.SYNOPSIS
A script to check for and attempt to repair common issues in the tools we use.

.DESCRIPTION
It will install the base applications we always want and will also uninstall the normal set as well as letting us do optional installed for Ops and Dev computers.
.EXAMPLE
minimumCoreFixes

.NOTES
notes

#>

#Patrick Moon - 2024
$global:errors=0;
function Invoke-Sanity-Checks {
    # Check if the script is running in PowerShell
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Output "This script must be run in PowerShell. Please open PowerShell ISE and run the script again."
        exit
    }
}

function ResetIECPL {
    Write-Host "We are going to attempt to reset the old IE settings..."
    Write-Host "A new popup is going to come up. You want to press the 'reset' button.  You do not need to check any boxes." -ForegroundColor Yellow
    RunDll32.exe InetCpl.cpl,ResetIEtoDefaults
}

function RunSFC {
    Write-Host "We are going to run SFC to fix any issues..."
    sfc /scannow
}

function RunSystemRepairFixes {
     Repair-WindowsImage -Online -CheckHealth
     DISM.exe /Online /Cleanup-image /RestoreHealth
}

function CheckSystemHealth {
    Write-Host "We are now going to check the system health of the OS..."
    $repairResult = Repair-WindowsImage -Online -CheckHealth
    if ($repairResult.ImageHealthState -ne "Healthy") {
        Write-Host "The system image is not healthy." -ForegroundColor Red
        $global:errors++
    } else {
        Write-Host "The system image is healthy." -ForegroundColor Green
    }
}
function CheckSystemStatus {
    Invoke-Sanity-Checks
    CheckWingetUpdate
    CheckSystemHealth
}
function RepairsToRun {
    ResetIECPL
    RunSFC
}

CheckSystemStatus
RepairsToRun