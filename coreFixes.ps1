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
#Patrick Moon - 2024
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # orginal: Start-Process -FilePath "powershell" -ArgumentList "-File .\coreSetup.ps1" -Verb RunAs
    # We are not running as administrator, so start a new process with 'RunAs'
    Start-Process powershell.exe "-File", ($myinvocation.MyCommand.Definition) -Verb RunAs
    exit
}

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
        Write-Host "Winget is either not installed or had an error. This is complicated. Good luck! Hint: check if App Installer is updated in the windows store." -ForegroundColor Red
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
    RunSystemRepairFixes
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

function RemoveAndBlockNewOutlook {
    # Path to the registry key
    $regPath = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\Orchestrator\UScheduler_Oobe"
    
    # Create the registry key if it doesn't exist
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force
    }
    
    # Set the registry value to block new Outlook
    $propertyName = "BlockedOobeUpdaters"
    $propertyValue = "MS_Outlook"
    
    try {
        # Attempt to set the property
        Set-ItemProperty -Path $regPath -Name $propertyName -Value $propertyValue
    } catch {
        # If the property doesn't exist, create it
        New-ItemProperty -Path $regPath -Name $propertyName -Value $propertyValue -PropertyType String -Force
    }
    
    # Remove the new Outlook app if it's already installed
    $outlookPackage = Get-AppxPackage -Name "Microsoft.OutlookForWindows"
    if ($outlookPackage) {
        Remove-AppxProvisionedPackage -AllUsers -Online -PackageName $outlookPackage.PackageFullName
    }
}
function FindProductCode {
    param (
        $applicationName = "Adobe Acrobat"
    )
    $products = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE (Name LIKE '%$applicationName%')"

    if ($products -is [System.Array]) {
        Write-Output "Found $($products.Count) instances of $applicationName. Please ensure only one instance is installed."
        return $null
    }

    return $products.IdentifyingNumber
}


function RepairAdobe {
    param (
        $productCode = "AC76BA86-7AD7-FFFF-7B44-AC0F074E4100"
    )

    $command = "msiexec /fa `"$productCode`""
    #Write-Output "Running command: $command"

    $process = Start-Process -FilePath "msiexec" -ArgumentList "/fa `"$productCode`"" -PassThru -Wait

    if ($process.ExitCode -eq 0) {
        Write-Output "Repair completed successfully."
    } else {
        Write-Output "Repair failed with exit code $($process.ExitCode)."
    }
}


CheckSystemStatus
#now we deal with errors
if($global:errors -gt 0) {
    Write-Host "We found an issue so we are starting repairs!" -ForegroundColor Red
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

Write-Host "Did you want to try repairing Adobe reader?"
$userInput = Read-Host " (N/y)"
if ($userInput -eq "y") {
    $productCode = FindProductCode
    RepairAdobe -productCode $productCode
}

Write-Host "Did you want to remove the new Outlook?"
$userInput = Read-Host " (N/y)"
if ($userInput -eq "y") {
    RemoveAndBlockNewOutlook
}

Write-Host "All operations you requested have completed."
Pause