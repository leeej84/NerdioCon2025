# Run as Administrator

# --- Function to create test users and assign to group ---
function Create-TestUsers {
    Write-Host "Creating local test users and group..." -ForegroundColor Cyan

    $Password = ConvertTo-SecureString "Password100" -AsPlainText -Force

    # Create local group 'TestUsers' if it doesn't exist
    if (-not (Get-LocalGroup -Name "TestUsers" -ErrorAction SilentlyContinue)) {
        New-LocalGroup -Name "TestUsers" -Description "Group for test users"
        Write-Host "Created group: TestUsers" -ForegroundColor Green
    }

    for ($i = 1; $i -le 20; $i++) {
        $username = "testuser{0:D2}" -f $i

        if (-not (Get-LocalUser -Name $username -ErrorAction SilentlyContinue)) {
            New-LocalUser -Name $username -Password $Password -FullName $username -Description "Test User $i" -UserMayNotChangePassword -PasswordNeverExpires
            Write-Host "Created user: $username" -ForegroundColor Green
        } else {
            Write-Warning "User $username already exists. Skipping."
        }

        # Add user to required groups
        Add-LocalGroupMember -Group "Remote Desktop Users" -Member $username -ErrorAction SilentlyContinue
        Add-LocalGroupMember -Group "TestUsers" -Member $username -ErrorAction SilentlyContinue
    }
}

# --- Ensure NuGet provider is installed silently ---
Write-Host "Installing NuGet provider silently..." -ForegroundColor Cyan
Install-PackageProvider -Name NuGet -Force -Scope AllUsers
Import-PackageProvider -Name NuGet -Force

# --- Install Selenium PowerShell module for all users ---
Write-Host "Installing Selenium PowerShell module for all users..." -ForegroundColor Cyan
Install-Module -Name Selenium -Force -Scope AllUsers

# --- Chocolatey install ---
if (!(Get-Command choco -ErrorAction SilentlyContinue)) {
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
}

# Refresh env vars
$env:Path += ";$($env:ChocolateyInstall)\bin"

# --- Install RDS Roles ---
Write-Host "Installing RDS roles..." -ForegroundColor Cyan

Install-WindowsFeature -Name RDS-RD-Server -IncludeManagementTools

# --- Install Office & Edge ---
Write-Host "Installing Edge and Office365..." -ForegroundColor Cyan

choco install microsoft-edge -y
choco install office365proplus -y --params "/SharedComputerLicensing:true"

# --- Confirm Office Shared Computer Activation Registry Key ---
$OfficeRegPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
if (Test-Path $OfficeRegPath) {
    Set-ItemProperty -Path $OfficeRegPath -Name SharedComputerLicensing -Value 1
    Write-Host "Office SCA registry key set." -ForegroundColor Green
} else {
    Write-Warning "Office registry path not found. Ensure Office 365 ProPlus is installed correctly."
}

# --- Create Users and Group ---
Create-TestUsers

# --- Download & Extract Automation Scripts ---
$zipUrl = "https://github.com/leeej84/NerdioCon2025/raw/refs/heads/main/Automation_Scripts/Automation_Scripts.zip"
$destFolder = "C:\Test_Scripts"
$zipFile = "$env:TEMP\Automation_Scripts.zip"

Write-Host "Downloading automation scripts..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile

if (!(Test-Path $destFolder)) {
    New-Item -Path $destFolder -ItemType Directory -Force
}

Write-Host "Extracting scripts to $destFolder..." -ForegroundColor Cyan
Expand-Archive -Path $zipFile -DestinationPath $destFolder -Force

# --- Confirm Manager.ps1 exists ---
$managerScript = Join-Path $destFolder "Manager.ps1"
$configPath = Join-Path $destFolder "Test_Config.json"

if (-not (Test-Path $managerScript)) {
    Write-Warning "Manager.ps1 not found in $destFolder. Make sure the downloaded archive contains it."
} else {
    Write-Host "Manager.ps1 ready at $managerScript" -ForegroundColor Green
}

Write-Host "Creating Startup shortcut to Manager.ps1 for all users..." -ForegroundColor Cyan

$startupFolder = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
$shortcutPath = Join-Path $startupFolder "Run Manager.lnk"
$targetPath = "powershell.exe"
$arguments = "-ExecutionPolicy Bypass -File `"$managerScript`" -ConfigPath `"$configPath`" -ScriptsPath `"$destFolder`""

# Create WScript.Shell COM object
$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetPath
$shortcut.Arguments = $arguments
$shortcut.WorkingDirectory = $destFolder
$shortcut.WindowStyle = 7 # Minimized
$shortcut.Save()

Write-Host "Shortcut created: $shortcutPath"

Write-Host "All setup complete! Test users will now auto-trigger Manager.ps1 with the required parameters at logon." -ForegroundColor Green

# --- Prompt for InfluxDB Token and Save to File ---
Write-Host "`nPlease enter the InfluxDB token. This will be saved securely to C:\Test_Scripts\Influx_Token.txt" -ForegroundColor Yellow
$token = Read-Host -AsSecureString "Influx Token"
$tokenPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($token)
)

$tokenPath = "C:\Test_Scripts\Influx_Token.txt"
$tokenPlain | Out-File -FilePath $tokenPath -Encoding UTF8 -Force

Write-Host "Influx token saved to $tokenPath" -ForegroundColor Green