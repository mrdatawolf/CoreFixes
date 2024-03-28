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
# Load assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$global:errors=0;
function Test-PSVersion {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        $logBox
    )
    $textBox.Text = "Checking powershell version"
    $logBox.AppendText("Checking powershell version`r`n")
    if ($PSVersionTable.PSVersion.Major -lt 5) {
       # $textBox.AppendText("Your powershell is too old!`r`n")
       $textBox.Text = "Your powershell is too old!"

        return $false
    }
    $logBox.AppendText("Powershell version is new enough!`r`n")

    return $true
}

function Test-Winget {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        $logBox
    )
    $logBox.AppendText("Checking winget`r`n")
    try {
        $wingetCheck = Get-Command winget -ErrorAction Stop
    }
    catch {
        $textBox.Text = "Winget has blocking issues!"
        $logBox.AppendText("Winget has blocking issues!`r`n")

        return $false
    }
    $logBox.AppendText("Winget seems to be working.`r`n")

    return $true
}

function Invoke-WingetUpdate {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        $logBox
    )
    $textBox.Text = "updating winget"
    $logBox.AppendText("Updating winget...`r`n")
    $output = & winget update 2>&1

    # Check if the output contains the error message
    if ($output -match "Failed in attempting to update the source: winget") {
        $textBox.Text = "Error: Failed attempting to update winget! Try updating 'App Installer'"
        $global:errors++
    } else {
        $logBox.AppendText("Winget update executed successfully.`r`n")
        $textBox.Text = "Done"
    }
}

function Invoke-ResetIECPL {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        $logBox
    )
    $logBox.AppendText("We are going to attempt to reset the old IE settings...`r`n")
    $logBox.AppendText("A new popup is going to come up. You want to press the 'reset' button.  You do not need to check any boxes.`r`n")
    RunDll32.exe InetCpl.cpl,ResetIEtoDefaults
    $textBox.Text = "Done"
    $logBox.AppendText("Done.`r`n")
}

function Invoke-SFC {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        $logBox
    )
    $logBox.AppendText("We are going to run SFC to fix any issues...")
    sfc /scannow
    $logBox.AppendText("Done")
}

function Invoke-SystemRepairFixes {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        $logBox
    )
    $logBox.AppendText("Repairing Windows Image")
    Repair-WindowsImage -Online -CheckHealth
    $logBox.AppendText("Running DISM")
    DISM.exe /Online /Cleanup-image /RestoreHealth
    $logBox.AppendText("Done")
}

function Invoke-SystemHealth {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        $logBox
    )
    $logBox.AppendText("We are now going to check the system health of the OS...")
    $repairResult = Repair-WindowsImage -Online -CheckHealth
    if ($repairResult.ImageHealthState -ne "Healthy") {
        $logBox.AppendText("The system image is not healthy.")
        $global:errors++
    } else {
        $logBox.AppendText("The system image is healthy.")
    }
}

function Invoke-RepairsToRun {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        $logBox
    )
    ResetIECPL
    RunSFC
}

function Invoke-Sara {
    param (
        $scenario,
        [System.Windows.Forms.TextBox]$textBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        $logBox
    )
    $logBox.AppendText("Checking for/creating SARA path")
    # Define the URL for the SARA tool
    $saraUrl = "https://aka.ms/SaRA_EnterpriseVersionFiles"
    # Define the local path where you want to save the tool
    $localPath = "~\Downloads\Sara"
    #check for Sara folder and create if needed.
    if (-not (Test-Path -Path $localPath)) {
        New-Item -Path $localPath -ItemType Directory -Force
    }
    $logBox.AppendText("Removing stale SARA files")
    #Now remove any files in the $localPath
    Get-ChildItem -Path $localPath | Remove-Item -Recurse -Force
    # Download the SARA tool
    $logBox.AppendText("Downloading a fresh copy")
    Invoke-WebRequest -Uri $saraUrl -OutFile "$localPath\SaRA.zip"
    # Extract the ZIP file
    Expand-Archive -Path "$localPath\SaRA.zip" -DestinationPath $localPath
    # Run the SARA tool with the desired scenario
    $logBox.AppendText("Running SARA")
    & "$localPath\SaraCmd.exe" -S $scenario -AcceptEula -CloseOffice
}

function Invoke-RepairOutlookO365 {
    param (
        [System.Windows.Forms.TextBox]$textBox,
        [System.Windows.Forms.ProgressBar]$progressBar,
        $logBox
    )
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
    $logBox.AppendText("Running Office repair")
    Start-Process -FilePath $OfficeClickToRunPath -ArgumentList $Arguments -NoNewWindow
}

#the following functions build the gui

function Initialize-Form {
    param(
        [hashtable] $ObjectDimensions = @{ Width=1024; Height=768 },
        [System.Windows.Forms.Form]$Form,
        [string]$Text = "My Form"
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Text
    $form.Size = New-Object System.Drawing.Size($ObjectDimensions.Width, $ObjectDimensions.Height)
    $form.StartPosition = "CenterScreen"
    return $form
}

function Initialize-Control {
    param(
        [hashtable] $ObjectDimensions = @{ Width=200; Height=20 },
        [PSCustomObject]$CurrentLocation = @{ X=50; Y=70 },
        [System.Windows.Forms.Form]$Form,
        [string]$ControlType,
        [string]$Text = ""
    )
    switch ($ControlType) {
        "ProgressBar" {
            $control = New-Object System.Windows.Forms.ProgressBar
            $control.Size = New-Object System.Drawing.Size($ObjectDimensions.Width, $ObjectDimensions.Height)
            $control.Value = 1
        }
        "TextBox" {
            $control = New-Object System.Windows.Forms.TextBox
            $control.Size = New-Object System.Drawing.Size($ObjectDimensions.Width, $ObjectDimensions.Height)
            $control.Readonly = $true
        }
        "SpinnerLabel" {
            $control = New-Object System.Windows.Forms.Label
            $control.Size = New-Object System.Drawing.Size($ObjectDimensions.Width, $ObjectDimensions.Height)
            $control.Text = $Text
        }
        "Log" {
            $control = New-Object System.Windows.Forms.TextBox
            $control.Size = New-Object System.Drawing.Size(200, 100)
            $control.Multiline = $true
            $control.ScrollBars = 'Vertical'
            $control.Readonly = $true   
        }
    }
    $control.Location = New-Object System.Drawing.Point($CurrentLocation.X, $CurrentLocation.Y)
    $CurrentLocation.X = 10
    $CurrentLocation.Y += $ObjectDimensions.Height

    return $control
}

function Initialize-PictureBox {
    param(
        [hashtable] $ObjectDimensions = @{ Width=60; Height=30 },
        [PSCustomObject]$CurrentLocation,
        [System.Windows.Forms.Form]$Form,
        [string]$Text,
        $url = "https://trustbiztech.com/public/logos/biztech.png"
    )
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $webClient = New-Object System.Net.WebClient
    $imagePath = [System.IO.Path]::GetTempFileName()
    $webClient.DownloadFile($url, $imagePath)
    $pictureBox.Image = [System.Drawing.Image]::Fromfile($imagePath)
    $pictureBox.Size = New-Object System.Drawing.Size($ObjectDimensions.Width, $ObjectDimensions.Height)  # Change this to your desired size
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $pictureBox.Location = New-Object System.Drawing.Point(($CurrentLocation.X - $pictureBox.Width - 20),($CurrentLocation.Y - $pictureBox.Height - 40)) 

    return $pictureBox
}

function Initialize-Button {
    param(
        [hashtable] $ObjectDimensions = @{ Width=150; Height=20 },
        [PSCustomObject]$CurrentLocation = @{ X=100; Y=35 },
        [System.Windows.Forms.Form]$Form,
        [string]$Text = "Click me",
        $NewLine = $true
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($CurrentLocation.X, $CurrentLocation.Y)
    $button.Size = New-Object System.Drawing.Size($ObjectDimensions.Width, $ObjectDimensions.Height)
    $button.Text = $Text
    
    if($NewLine) {
    $CurrentLocation.X = 10
    $CurrentLocation.Y += $ObjectDimensions.Height
    } else {
        $CurrentLocation.X += $ObjectDimensions.Width
    }
    
    return $button
}

function Invoke-Spinner {
    param(
        $spinnerLabel,
        $currentStep
    )
    $spinnerChars = @('/', '|', '\', '-')
    $spinnerLabel.Text = $spinnerChars[$currentStep % $spinnerChars.Length]
}

function Invoke-ProgressBar {
    param(
        $progressBar,
        $currentStep
    )
    $progressBar.Value = $currentStep
}
$location = New-Object -TypeName psobject -Property @{ X=10; Y=10 }
$form = Initialize-Form -CurrentLocation $location
$sanityButton = Initialize-Button -CurrentLocation $location -Text "Sanity Checks" -NewLine $false 
$winGetUpdateButton = Initialize-Button -CurrentLocation $location -Text "Update Winget" -NewLine $false
$systemHealthButton = Initialize-Button -CurrentLocation $location -Text "Check System Health" -NewLine $false
$systemRepairButton = Initialize-Button -CurrentLocation $location -Text "Repair System Health" -NewLine $false
$saraButton = Initialize-Button -CurrentLocation $location -Text "SARA office activation" -NewLine $false
$o365Button = Initialize-Button -CurrentLocation $location -Text "o365 repair" -NewLine $true
$overallProgressBar = Initialize-Control -ControlType "ProgressBar" -CurrentLocation $location -Form $form
$progressBar = Initialize-Control -ControlType "ProgressBar" -CurrentLocation $location -Form $form
$textBox = Initialize-Control -ControlType "TextBox" -CurrentLocation $location -Form $form
$logBox = Initialize-Control -ControlType "Log" -CurrentLocation $location -Form $form 
#special ui
$location.X = 10
$location.Y = $form.Height-60
$spinnerLabel = Initialize-Control -ControlType "SpinnerLabel" -CurrentLocation $location -Form $form -Text "-"
$location.X = $form.Width
$location.Y = $form.Height
$pictureBox = Initialize-PictureBox -CurrentLocation $location



#now we add the commands for the buttons
$sanityButton.Add_Click({
    $progressBar.Value = 0
    Invoke-ProgressBar -progressBar $progressBar -textBox $textBox -currentStep $i
    Invoke-Spinner -spinnerLabel $spinnerLabel -currentStep $i
    $logBox.AppendText("Initial sanity checks.`r`n")
    if (Test-PSVersion -textBox $textBox -logBox $logBox) {
        $textBox.Text = "PS version is good enough"
    } else {
        exit
    }
    $progressBar.Value = 50
    if (Test-Winget -textBox $textBox -logBox $logBox) {
        $textBox.Text = "Winget seems to be responding"
    } else {
        exit
    }
    $textBox.Text = "Done"
    $logBox.AppendText("Sanity checks done.`r`n")
    $progressBar.Value = 100
    $overallProgressBar.Value += 16
    $this.Enabled = $false
}) 
$winGetUpdateButton.Add_Click({
    $progressBar.Value = 0
    Invoke-ProgressBar -progressBar $progressBar -textBox $textBox -currentStep $i
    Invoke-Spinner -spinnerLabel $spinnerLabel -currentStep $i
    Invoke-WingetUpdate -textBox $textBox -logBox $logBox -progressBar $progressBar
    $progressBar.Value = 100
    $overallProgressBar.Value += 16
    $this.Enabled = $false
}) 
$systemHealthButton.Add_Click({
    $progressBar.Value = 0
    Invoke-ProgressBar -progressBar $progressBar -textBox $textBox -currentStep $i
    Invoke-Spinner -spinnerLabel $spinnerLabel -currentStep $i
    Invoke-SystemHealth -textBox $textBox -logBox $logBox -progressBar $progressBar
    #now we deal with errors
    if($global:errors -gt 0) {
        Write-Host "We found an issue so we are starting repairs!" -ForegroundColor Green
        Invoke-RepairsToRun -textBox $textBox -logBox $logBox -progressBar $progressBar
    }
    $progressBar.Value = 100
    $overallProgressBar.Value += 16
    $this.Enabled = $false
})
$systemRepairButton.Add_Click({
    $progressBar.Value = 0
    Invoke-ProgressBar -progressBar $progressBar -textBox $textBox -currentStep $i
    Invoke-Spinner -spinnerLabel $spinnerLabel -currentStep $i
    Invoke-RepairsToRun -textBox $textBox -logBox $logBox -progressBar $progressBar
    $progressBar.Value = 100
    $overallProgressBar.Value += 16
    $this.Enabled = $false
}) 
$saraButton.Add_Click({
    $progressBar.Value = 0
    Invoke-ProgressBar -progressBar $progressBar -textBox $textBox -currentStep $i
    Invoke-Spinner -spinnerLabel $spinnerLabel -currentStep $i
    Invoke-Sara -scenario "ResetOfficeActivation" -textBox $textBox -logBox $logBox -progressBar $progressBar
    $progressBar.Value = 100
    $overallProgressBar.Value += 16
    $this.Enabled = $false
}) 
$o365Button.Add_Click({
    $progressBar.Value = 0
    Invoke-ProgressBar -progressBar $progressBar -textBox $textBox -currentStep $i
    Invoke-Spinner -spinnerLabel $spinnerLabel -currentStep $i
    Invoke-RepairOutlookO365 -textBox $textBox -logBox $logBox -progressBar $progressBar
    $progressBar.Value = 100
    $overallProgressBar.Value += 16
    $this.Enabled = $false
}) 

$form.Controls.Add($sanityButton)
$form.Controls.Add($winGetUpdateButton)
$form.Controls.Add($systemHealthButton)
$form.Controls.Add($systemRepairButton)
$form.Controls.Add($saraButton)
$form.Controls.Add($o365Button)
$form.Controls.Add($overallProgressBar)
$form.Controls.Add($progressBar)
$form.Controls.Add($textBox)
$form.Controls.Add($logBox)
$form.Controls.Add($spinnerLabel)
$form.Controls.Add($pictureBox)

$form.ShowDialog()