# Run as Administrator

# --- Function to create test users ---
function Create-TestUsers {
    Write-Host "Creating local test users..." -ForegroundColor Cyan

    $Password = ConvertTo-SecureString "Password100" -AsPlainText -Force

    for ($i = 1; $i -le 10; $i++) {
        $username = "testuser{0:D2}" -f $i

        if (-not (Get-LocalUser -Name $username -ErrorAction SilentlyContinue)) {
            New-LocalUser -Name $username -Password $Password -FullName $username -Description "Test User $i" -UserMayNotChangePassword -PasswordNeverExpires
            Add-LocalGroupMember -Group "Remote Desktop Users" -Member $username
            Write-Host "Created user: $username" -ForegroundColor Green
        } else {
            Write-Warning "User $username already exists. Skipping."
        }
    }
}

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

# --- Create Test Users ---
Create-TestUsers

Write-Host "`nAll steps completed successfully. You may need to reboot to finalize RDS setup." -ForegroundColor Green
