<#
.SYNOPSIS
A script to check for and attempt to repair common issues in the tools we use.

.DESCRIPTION
It will install the base applications we always want and will also uninstall the normal set as well as letting us do optional installed for Ops and Dev computers.
.EXAMPLE
coreFixes

.NOTES
notes

#>
#Requires -RunAsAdministrator
#Patrick Moon - 2024
$global:errors=0;
function Invoke-Sanity-Checks {
    # Check if the script is running in PowerShell
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Output "This script must be run in PowerShell. Please open PowerShell and run the script again."
        exit
    }

    # Check if winget is installed
    try {
        $wingetCheck = Get-Command winget -ErrorAction Stop
        Write-Host "Winget is installed so we can continue." -ForegroundColor Green
    } catch {
        Write-Host "Winget is not installed/ had an error. This is complicated. Good luck!" -ForegroundColor Red
        exit
    }
}

function CheckWingetUpdate {
    Write-Host "We are going to check if winget is able to update its self".
    $output = & winget update 2>&1

    # Check if the output contains the error message
    if ($output -match "Failed in attempting to update the source: winget") {
        Write-Host "Error: Failed attempting to update winget! Try updating 'App Installer'" -ForegroundColor Red
        $global:errors++
    } else {
        Write-Host "Winget update executed successfully." -ForegroundColor Green
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

function DoSaraWork($scenario) {
    # Define the URL for the SARA tool
    $saraUrl = "https://aka.ms/SaRA_EnterpriseVersionFiles"

    # Define the local path where you want to save the tool
    $localPath = "~\Downloads\Sara"

    #check for Sara folder and create if needed.
    if (-not (Test-Path -Path $localPath)) {
        New-Item -Path $localPath -ItemType Directory -Force
    }

    #Now remove any files in the $localPath
    Get-ChildItem -Path $localPath | Remove-Item -Recurse -Force

    # Download the SARA tool
    Invoke-WebRequest -Uri $saraUrl -OutFile "$localPath\SaRA.zip"

    # Extract the ZIP file
    Expand-Archive -Path "$localPath\SaRA.zip" -DestinationPath $localPath

    # Run the SARA tool with the desired scenario
    & "$localPath\SaraCmd.exe" -S $scenario -AcceptEula -CloseOffice
}

function RepairOutlookO365 {
    param (
        $RepairScenario = "Repair"
    )
    # Path to OfficeClickToRun.exe (change accordingly if your path is different)
    $OfficeClickToRunPath = "C:\\Program Files\\Microsoft Office 15\\ClientX64\\OfficeClickToRun.exe"

    # Platform (x64 or x86)
    $Platform = "x64" # change to x86 if you're using 32-bit Office

    # Language culture
    $Culture = "en-us" # change to your language culture

    # Command to start the repair
    $Arguments = "scenario=$RepairScenario", "platform=$Platform", "culture=$Culture", "DisplayLevel=True"

    # Run the command
    Start-Process -FilePath $OfficeClickToRunPath -ArgumentList $Arguments -NoNewWindow
}

CheckSystemStatus
#now we deal with errors
if($global:errors -gt 0) {
    Write-Host "We found an issue so we are starting repairs!" -ForegroundColor Green
    RepairsToRun
    
} else {
    #if no errors were found we should still offer to run fixes.
    Write-Host "No issues were detected." -ForegroundColor Green
    Write-Host "Did you still want to try running our common fixes?"
    $userInput = Read-Host " (n/Y)" 
    if ($userInput -eq "n") {
        Write-Host "No problem, we will skip running those repairs!" -ForegroundColor Green
    } else {
        RepairsToRun
    }
}

Write-Host "Did you want Sara to reset office activation?"
$userInput = Read-Host " (N/y)" 
if ($userInput -eq "y") {
    DoSaraWork("ResetOfficeActivation")
}

Write-Host "Did you want to run the office365 repair?"
$userInput = Read-Host " (N/y)" 
if ($userInput -eq "y") {
    RepairOutlookO365
}