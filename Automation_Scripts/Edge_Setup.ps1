# Define Selenium and WebDriver paths
$seleniumPath = "C:\Selenium"
$edgeDriverPath = "$seleniumPath\msedgedriver.exe"

# Ensure Selenium directory exists
if (!(Test-Path $seleniumPath)) {
    New-Item -ItemType Directory -Path $seleniumPath -Force | Out-Null
}

# Detect PowerShell Version
$psVersion = $PSVersionTable.PSVersion.Major

# Install NuGet Provider if missing
if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
    Write-Host "Installing NuGet Provider..."
    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
}

# Install Selenium Module via PowerShellGet
if (-not (Get-Module -ListAvailable -Name Selenium)) {
    Write-Host "Installing Selenium module..."
    Install-Module -Name Selenium -Force -Scope CurrentUser
}

# Verify Selenium installation
if (-not (Get-Module -ListAvailable -Name Selenium)) {
    Write-Host "Selenium module failed to install. Exiting..."
    exit 1
}

# Function to find WebDriver.dll dynamically for both PowerShell 5 and 7
function Find-SeleniumDLL {
    $possiblePaths = @(
        "$env:USERPROFILE\OneDrive\Documents\WindowsPowerShell\Modules\Selenium",
        "$env:USERPROFILE\OneDrive\Documents\PowerShell\Modules\Selenium",
        "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\Selenium",
        "$env:USERPROFILE\Documents\PowerShell\Modules\Selenium",
        "$env:ProgramFiles\WindowsPowerShell\Modules\Selenium",
        "$env:ProgramFiles\PackageManagement\NuGet\Packages\Selenium.WebDriver"
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $dllPath = Get-ChildItem -Path $path -Recurse -Filter "WebDriver.dll" | Select-Object -ExpandProperty FullName -First 1
            if ($dllPath) { return $dllPath }
        }
    }
    
    Write-Host "WebDriver.dll not found in the Selenium module. Exiting..."
    exit 1
}

# Get the installed Edge version
function Get-EdgeVersion {
    $edgePaths = @(
        "HKLM:\SOFTWARE\Microsoft\Edge\BLBeacon",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon",
        "HKCU:\SOFTWARE\Microsoft\Edge\BLBeacon"
    )

    foreach ($path in $edgePaths) {
        if (Test-Path $path) {
            return (Get-ItemProperty -Path $path -Name version).version
        }
    }

    Write-Host "Microsoft Edge is not installed or not found in the registry."
    exit 1
}

# Download the correct Edge WebDriver
function Download-EdgeWebDriver {
    param ([string]$edgeVersion)

    $baseUrl = "https://msedgedriver.azureedge.net/$edgeVersion/edgedriver_win64.zip"
    $zipPath = "$seleniumPath\edgedriver.zip"

    Write-Host "Downloading Edge WebDriver for version $edgeVersion..."
    Invoke-WebRequest -Uri $baseUrl -OutFile $zipPath

    Write-Host "Extracting WebDriver..."
    Expand-Archive -Path $zipPath -DestinationPath $seleniumPath -Force

    # Ensure the correct filename
    $driverPath = Get-ChildItem -Path $seleniumPath -Filter "msedgedriver.exe" -Recurse | Select-Object -ExpandProperty FullName -First 1
    if ($driverPath -and $driverPath -ne $edgeDriverPath) {
        Move-Item -Path $driverPath -Destination $edgeDriverPath -Force
    }

    Remove-Item $zipPath -Force
}

# Check if Edge WebDriver exists, if not, download it
if (!(Test-Path $edgeDriverPath)) {
    $edgeVersion = Get-EdgeVersion
    Download-EdgeWebDriver -edgeVersion $edgeVersion
}

# Load Selenium Assembly dynamically
$seleniumDllPath = Find-SeleniumDLL
Add-Type -Path $seleniumDllPath

# Set up Edge WebDriver service (Fix: Explicitly specify driver path)
$env:Path += ";$seleniumPath"  # Ensure WebDriver is in system path
$edgeService = [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService($seleniumPath, "msedgedriver.exe")
$edgeService.HideCommandPromptWindow = $true

# Set up EdgeOptions correctly
$edgeOptions = New-Object OpenQA.Selenium.Edge.EdgeOptions
if ($psVersion -ge 7) {
    # PowerShell 7 uses AddAdditionalCapability
    $edgeOptions.AddAdditionalCapability("ms:edgeOptions", @{"args" = @("--start-maximized")})
} else {
    # PowerShell 5 uses AddArgument
    $edgeOptions.AddArgument("--start-maximized")
}

# Start Edge WebDriver
Write-Host "Launching Edge Browser..."
$driver = New-Object OpenQA.Selenium.Edge.EdgeDriver($edgeService, $edgeOptions)

# Navigate to Google
Write-Host "Navigating to Google..."
$driver.Navigate().GoToUrl("https://www.google.com")

# Close browser
Write-Host "Closing Edge..."
$driver.Quit()
