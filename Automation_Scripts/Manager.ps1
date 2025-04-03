param (
    [Parameter()]
    [string]$ConfigPath = ".\Test_Config.json",

    [Parameter()]
    [string]$ScriptsPath = ".\",

    [Parameter()]
    [string]$influxServer = "INFLUX01"
)

# Set common tags
$commonTags = @{ location = "London"; machine = $env:COMPUTERNAME; username = $env:USERNAME }

# Load config
if (-not (Test-Path $ConfigPath)) {
    Write-Error "Config file not found at $ConfigPath"
    exit 1
}

$config = Get-Content $ConfigPath | ConvertFrom-Json
$loops = $config.Loops
$durationMinutes = $config.Time_In_Minutes

# Log timing
$startTime = Get-Date

function Run-Script {
    param (
        [string]$ScriptName
    )
    $scriptPath = Join-Path $ScriptsPath $ScriptName
    if (Test-Path $scriptPath) {
        Write-Host "Executing: $ScriptName"
        & $scriptPath -Tags ($commonTags + @{ script = "$ScriptName" })
        Start-Sleep -Seconds 5
    } else {
        Write-Warning "Script not found: $scriptPath"
    }
}

# --- Launch and Close Edge to ensure registry values are populated ---
Write-Host "Launching and closing Edge to populate registry..." -ForegroundColor Cyan
Start-Process "msedge.exe" -ArgumentList "--no-first-run" -WindowStyle Minimized
Start-Sleep -Seconds 5
Get-Process "msedge" -ErrorAction SilentlyContinue | Stop-Process -Force
Write-Host "Edge has been launched and closed." -ForegroundColor Green

# Determine loop mode
if ($loops -gt 0) {
    Write-Host "Running $loops loop(s) of automation..."
    for ($i = 1; $i -le $loops; $i++) {
        Write-Host "`n--- Loop $i ---"
        Run-Script 'Edge_Setup.ps1'
        Run-Script 'Edge_BBC.ps1'
        Run-Script 'Excel.ps1'
        Run-Script 'PowerPoint.ps1'
        Run-Script 'Word.ps1'
    }
}
elseif ($durationMinutes -gt 0) {
    Write-Host "Running automation for $durationMinutes minute(s)..."
    do {
        Run-Script 'Edge_Setup.ps1'
        Run-Script 'Edge_BBC.ps1'
        Run-Script 'Excel.ps1'
        Run-Script 'PowerPoint.ps1'
        Run-Script 'Word.ps1'
    } while ((Get-Date) -lt $startTime.AddMinutes($durationMinutes))
}
else {
    Write-Warning "No valid loop or time configuration found. Exiting."
    exit 1
}

Write-Host "`nAll tasks completed at $(Get-Date)."
