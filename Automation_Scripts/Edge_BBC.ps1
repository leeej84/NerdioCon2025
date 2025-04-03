#Define parameters
param (
    [hashtable]$Tags = @{}
)

# Define Selenium and WebDriver paths
$seleniumPath = "C:\Selenium"
$edgeDriverPath = "$seleniumPath\msedgedriver.exe"

# Function to send timing data to Influx
function Send-TimingToInflux {
    param (
        [string]$taskName,
        [int]$durationMs,
        [hashtable]$Tags
    )

    if (-not ($durationMs -is [int]) -or $durationMs -lt 0) {
        Write-Host "Invalid duration detected for '$taskName'. Skipping InfluxDB write."
        return
    }

    $influxURL = "http://influxdb-01.ctxlab.local:8086/api/v2/write?org=Performance&bucket=Performance&precision=ms"
    $influxToken = "$(Get-Content .\Influx_Token.txt)"
    $timestamp = [math]::Round((Get-Date).ToUniversalTime().Subtract([datetime]"1970-01-01").TotalMilliseconds)

    # Compose tag string
    $tagParts = @()
    foreach ($key in $Tags.Keys) {
        $tagParts += "$($key)=$($Tags[$key] -replace ' ', '\ ')"
    }
    $tagString = $tagParts -join ','

    $taskName = $taskName -replace ' ', '\ '
    $measurement = "automation_timings,task=$taskName"

    if ($tagString -ne "") {
        $measurement += ",$tagString"
    }

    $body = "$measurement duration=$durationMs $timestamp"

    try {
        Invoke-RestMethod -Uri $influxURL -Method Post -Headers @{ "Authorization" = "Token $influxToken" } -Body $body
        Write-Host "Sent timing for '$taskName' ($durationMs ms) with tags: $tagString"
    } catch {
        Write-Host "Failed to send timing data: $_"
    }
}



# Function to measure execution time
function Measure-ExecutionTime {
    param (
        [string]$ActionName,
        [scriptblock]$Action,
        [hashtable]$Tags
    )
    try {
        $startTime = Get-Date
        $result = Invoke-Command -ScriptBlock $Action
        $endTime = Get-Date
        $elapsed = ($endTime - $startTime).TotalMilliseconds
        $null = Send-TimingToInflux -taskName $ActionName -durationMs $elapsed -Tags $Tags
        Write-Host "$ActionName took $($elapsed) ms"
        return $result
    } catch {
        Write-Host "Error during $ActionName : $_"
        return $null
    }
}

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

# Ensure Selenium directory exists
if (!(Test-Path $seleniumPath)) {
    New-Item -ItemType Directory -Path $seleniumPath -Force | Out-Null
}

# Load Selenium Assembly dynamically
$seleniumDllPath = Find-SeleniumDLL
Add-Type -Path $seleniumDllPath

# Set up Edge WebDriver service
$env:Path += ";$seleniumPath"
$edgeService = Measure-ExecutionTime -ActionName "Create Edge WebDriver Service" -Action {
    [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService($seleniumPath, "msedgedriver.exe")
} -Tags $Tags
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
$driver = Measure-ExecutionTime -ActionName "Launch Edge Browser" -Action {
    New-Object OpenQA.Selenium.Edge.EdgeDriver($edgeService, $edgeOptions)
} -Tags $Tags

# Navigate to BBC
Measure-ExecutionTime -ActionName "Navigate to BBC" -Action {
    $driver.Navigate().GoToUrl("https://www.bbc.com")
} -Tags $Tags

Start-Sleep -Seconds 5  # Wait for elements to load

# Try finding the 'Reject All' button
Measure-ExecutionTime -ActionName "Click 'Reject All' Button" -Action {
    try {
        $rejectButton = $driver.FindElement([OpenQA.Selenium.By]::XPath("//*[@id='header-content']/section/div/div/div[2]/button[2]"))
        $rejectButton.Click()
        Write-Host "Clicked 'Reject All' button on BBC."
    } catch {
        Write-Host "Could not find the 'Reject All' button."
    }
} -Tags $Tags

# Function for Smooth Scrolling
function Smooth-Scroll {
    Measure-ExecutionTime -ActionName "Smooth Scroll" -Action {
        $driver.ExecuteScript(@"
        var scrollInterval = setInterval(function() {
            window.scrollBy(0, 10); // Scroll down 10 pixels at a time
            if ((window.innerHeight + window.scrollY) >= document.body.scrollHeight) {
                clearInterval(scrollInterval); // Stop when reaching the bottom
            }
        }, 50);
"@)
    } -Tags $Tags
}

# Navigate through some tabs
$tabs = @{
    "News" = "//*[@id='global-navigation']/div[2]/ul[2]/li[2]/a/span"  # News
    "Sport" = "//*[@id='global-navigation']/div[2]/ul[2]/li[3]/a/span"  # Sport
    "Weather" = "//*[@id='global-navigation']/div[2]/ul[2]/li[4]/a/span"  # Weather
    "Sounds" = "//*[@id='orb-header']/div/nav[1]/ul/li[6]/a"  # Sounds
}

foreach ($tabName in $tabs.Keys) {
    $xpath = $tabs[$tabName]
    
    try {
        Write-Host "Navigating to $tabName tab..."
        $tabElement = $driver.FindElement([OpenQA.Selenium.By]::XPath($xpath))
        $tabElement.Click()
        Smooth-Scroll
        Start-Sleep -Seconds 15
        Write-Host "Successfully navigated to $tabName tab."
    } catch {
        Write-Host "Failed to navigate to $tabName tab. XPath: $xpath"
    }
}

# Wait before closing
Start-Sleep -Seconds 5
Measure-ExecutionTime -ActionName "Quit Browser" -Action {
    $driver.Quit()
} -Tags $Tags
