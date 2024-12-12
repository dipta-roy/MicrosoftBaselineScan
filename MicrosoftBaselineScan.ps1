###################################################
#
#	Title 	: Microsoft Baseline Scanner
#	Author	: Dipta Roy
#	Version	: 1.0
#	Date	: 12-12-2024
#	Usage	: Run cmd as administrator and run PowerShell -ExecutionPolicy Bypass .\MicrosoftBaselineScan.ps1
#	CAB file: Download latest wsusscn2.cab from https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab
#
###################################################

# Ensure the script is run as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ERROR: This script must be run as Administrator. Run the program as Administrator..."
	Exit
}

# Prompt user to select the CAB file
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Filter = "CAB files (*.cab)|*.cab"
$OpenFileDialog.Title = "Select the WSUS offline scan CAB file"
if ($OpenFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No file selected. Exiting."
    Exit
}

$CabFilePath = $OpenFileDialog.FileName

# Validate the selected CAB file
if (-not (Test-Path $CabFilePath)) {
    Write-Host "The selected file does not exist. Exiting."
    Exit
}

# Define output file path in the script's directory
$ScriptDirectory = $PSScriptRoot
$ComputerName = $env:COMPUTERNAME
$CurrentDateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutputFileName = Join-Path $ScriptDirectory "$ComputerName`_$CurrentDateTime.csv"

# Initialize COM objects for Windows Update scanning
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
$UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $CabFilePath)
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

Write-Host "Searching for updates..."
$UpdateSearcher.ServerSelection = 3  # ssOthers
$UpdateSearcher.ServiceID = [string]$UpdateService.ServiceID

# Perform the search
$SearchResult = $UpdateSearcher.Search("IsInstalled=0")
$Updates = $SearchResult.Updates

if ($SearchResult.Updates.Count -eq 0) {
    Write-Host "There are no applicable updates."
    Exit
}

# Save results to a CSV file
$Results = @()
For ($i = 0; $i -lt $SearchResult.Updates.Count; $i++) {
    $update = $SearchResult.Updates.Item($i)
    $Results += [PSCustomObject]@{
        Index       = $i + 1
        Title       = $update.Title
        Description = $update.Description
    }
}

try {
    $Results | Export-Csv -Path $OutputFileName -NoTypeInformation -Encoding UTF8
    Write-Host "Update scan completed. Results saved to: $OutputFileName"
} catch {
    Write-Host "Failed to save results to $OutputFileName. Error: $_"
}