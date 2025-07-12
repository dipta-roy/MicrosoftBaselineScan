###################################################
#
#   Title   : Microsoft Baseline Scanner
#   Author  : Dipta Roy
#   Version : 2.0
#   Date    : 13-07-2025
#   Usage   : Run as Administrator
#             PowerShell -ExecutionPolicy Bypass .\MicrosoftBaselineScan.ps1
#   Note    : Download latest wsusscn2.cab from:
#             https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab
#
###################################################

# --- Step 1: Check for Administrator Privileges ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "ERROR: Script must be run as Administrator."
    exit 1
}

# --- Step 2: Prompt user to select a CAB file ---
Add-Type -AssemblyName System.Windows.Forms

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Filter = "CAB files (*.cab)|*.cab"
$OpenFileDialog.Title  = "Select WSUS Offline Scan CAB File"

if ($OpenFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No file selected. Exiting..."
    exit 1
}

$CabFilePath = $OpenFileDialog.FileName

if (-not (Test-Path $CabFilePath)) {
    Write-Error "The selected CAB file does not exist."
    exit 1
}

# --- Step 3: Prepare output file path ---
$ScriptDirectory   = $PSScriptRoot
$ComputerName      = $env:COMPUTERNAME
$CurrentDateTime   = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutputFileName    = Join-Path $ScriptDirectory "$ComputerName`_$CurrentDateTime.csv"

# --- Step 4: Set up COM objects for update scan ---
try {
    $UpdateSession       = New-Object -ComObject Microsoft.Update.Session
    $UpdateServiceMgr    = New-Object -ComObject Microsoft.Update.ServiceManager
    $UpdateService       = $UpdateServiceMgr.AddScanPackageService("Offline Sync Service", $CabFilePath)
    $UpdateSearcher      = $UpdateSession.CreateUpdateSearcher()

    $UpdateSearcher.ServerSelection = 3  # ssOthers
    $UpdateSearcher.ServiceID       = [string]$UpdateService.ServiceID
} catch {
    Write-Error "Failed to initialize update services: $_"
    exit 1
}

# --- Step 5: Search for missing updates ---
Write-Host "Searching for applicable updates..."

try {
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
} catch {
    Write-Error "Update search failed: $_"
    exit 1
}

if ($SearchResult.Updates.Count -eq 0) {
    Write-Host "No applicable updates found."
    exit 0
}

# --- Step 6: Extract update info into CSV format ---
$Results = for ($i = 0; $i -lt $SearchResult.Updates.Count; $i++) {
    $update = $SearchResult.Updates.Item($i)
    [PSCustomObject]@{
        Index       = $i + 1
        Title       = $update.Title
        Description = $update.Description
    }
}

# --- Step 7: Save results ---
try {
    $Results | Export-Csv -Path $OutputFileName -NoTypeInformation -Encoding UTF8
    Write-Host "Scan completed. Results saved to: $OutputFileName"
} catch {
    Write-Error "Failed to save CSV file: $_"
    exit 1
}
