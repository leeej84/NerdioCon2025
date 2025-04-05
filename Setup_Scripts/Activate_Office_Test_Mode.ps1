# Define the Office installation path and OSPP.VBS script path
$officePaths = @(
    "${env:ProgramFiles}\Microsoft Office\Office16\OSPP.VBS",
    "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"
)

$osppPath = $null

foreach ($path in $officePaths) {
    if (Test-Path $path) {
        $osppPath = $path
        break
    }
}

if (-not $osppPath) {
    Write-Error "OSPP.VBS script not found. Ensure Office 365 is installed."
    exit 1
}

# --- Enable Shared Computer Licensing ---
$clickToRunPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
if (-not (Test-Path $clickToRunPath)) {
    New-Item -Path $clickToRunPath -Force | Out-Null
}
Set-ItemProperty -Path $clickToRunPath -Name "SharedComputerLicensing" -Value 1 -Type String
Write-Host "Shared Computer Licensing enabled."

# --- Install the 10-day subscription key ---
$subscriptionKey = "DRNV7-VGMM2-B3G9T-4BF84-VMFTK"
Write-Host "Installing 10-day subscription key..."
cscript.exe $osppPath /inpkey:$subscriptionKey

# --- Disable Activation UI ---
$licensingPath = "HKLM:\Software\Microsoft\Office\16.0\Common\Licensing"
if (-not (Test-Path $licensingPath)) {
    New-Item -Path $licensingPath -Force | Out-Null
}
Set-ItemProperty -Path $licensingPath -Name "DisableActivationUI" -Value 1 -Type String
Write-Host "Activation UI disabled."

# --- Accept EULA ---
$registrationPath = "HKCU:\Software\Microsoft\Office\16.0\Registration"
if (-not (Test-Path $registrationPath)) {
    New-Item -Path $registrationPath -Force | Out-Null
}
Set-ItemProperty -Path $registrationPath -Name "AcceptAllEulas" -Value 1 -Type String
Write-Host "EULA automatically accepted."

# --- Check License Status ---
Write-Host "Checking license status..."
cscript.exe $osppPath /dstatus

# --- Optionally Rearm License ---
$rearm = Read-Host "Do you want to rearm the license to extend the grace period? (Y/N)"
if ($rearm -eq 'Y' -or $rearm -eq 'y') {
    Write-Host "Rearming license..."
    cscript.exe $osppPath /rearm
    Write-Host "License rearmed. Grace period extended."
}

Write-Host "Office 365 activation (with Shared Computer Licensing) completed."
