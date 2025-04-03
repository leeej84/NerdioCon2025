#Define parameters
param (
    [hashtable]$Tags = @{}
)

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

# Import WinAPI for window manipulation
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WinAPI {
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    }
"@

# Start overall timer
$overallStart = Get-Date

# Initialize
$ppt = $null
$presentation = $null
$slideShow = $null

try {
    $ppt = Measure-ExecutionTime -ActionName "Start PowerPoint Application" -Action {
        $app = New-Object -ComObject PowerPoint.Application
        $app.Visible = 1
        return $app
    } -Tags $Tags

    if ($null -eq $ppt) { throw "Failed to start PowerPoint application." }

    $presentation = Measure-ExecutionTime -ActionName "Create Presentation" -Action {
        $ppt.Presentations.Add()
    } -Tags $Tags

    if ($null -eq $presentation) { throw "Failed to create a new presentation." }

    Start-Sleep -Seconds 1
    $hWnd = [WinAPI]::FindWindow("PPTFrameClass", $ppt.Caption)
    if ($hWnd -ne [IntPtr]::Zero) {
        [WinAPI]::MoveWindow($hWnd, 100, 100, 1024, 768, $true) | Out-Null
    } else {
        Write-Host "Failed to find PowerPoint window handle."
    }

    Measure-ExecutionTime -ActionName "Ensure PowerPoint is Active" -Action {
        $shell = New-Object -ComObject WScript.Shell
        $shell.AppActivate($ppt.Caption) | Out-Null
    } -Tags $Tags

    # Add 3 slides
    $slide1 = Measure-ExecutionTime -ActionName "Add Slide 1" -Action {
        $presentation.Slides.Add(1, 1)
    } -Tags $Tags
    if ($slide1) {
        $slide1.Shapes.Title.TextFrame.TextRange.Text = "Slide 1: Introduction"
        $slide1.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Welcome to the automated PowerPoint presentation."
    }

    $slide2 = Measure-ExecutionTime -ActionName "Add Slide 2" -Action {
        $presentation.Slides.Add(2, 1)
    } -Tags $Tags
    if ($slide2) {
        $slide2.Shapes.Title.TextFrame.TextRange.Text = "Slide 2: Content"
        $slide2.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "This slide contains some important information."
    }

    $slide3 = Measure-ExecutionTime -ActionName "Add Slide 3" -Action {
        $presentation.Slides.Add(3, 1)
    } -Tags $Tags
    if ($slide3) {
        $slide3.Shapes.Title.TextFrame.TextRange.Text = "Slide 3: Conclusion"
        $slide3.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Thank you for watching this automated demo."
    }

    Start-Sleep -Seconds 2

    $savePath = "$env:TEMP\TestPresentation.pptx"
    if ($presentation) {
        Measure-ExecutionTime -ActionName "Save Presentation" -Action {
            $presentation.SaveAs($savePath)
        } -Tags $Tags
    } else {
        Write-Host "Skipping Save: Presentation object is null!"
    }

    Start-Sleep -Seconds 2

    Measure-ExecutionTime -ActionName "Start Slideshow" -Action {
        $slideShow = $presentation.SlideShowSettings
        $slideShow.AdvanceMode = 2
        $slideShow.StartingSlide = 1
        $slideShow.EndingSlide = $presentation.Slides.Count
        $slideShow.Run()
    } -Tags $Tags

    Start-Sleep -Seconds 7
    Write-Host "Automation completed successfully. Presentation saved at: $savePath"

} catch {
    Write-Host "An error occurred: $_"
} finally {
    if ($presentation.SlideShowWindow -ne $null) {
        try {
            Measure-ExecutionTime -ActionName "Stop Slideshow" -Action {
                $presentation.SlideShowWindow.View.Exit()
            } -Tags $Tags
        } catch {
            Write-Host "Failed to stop slideshow: $_"
        }
    }

    if ($presentation -ne $null) {
        try {
            Measure-ExecutionTime -ActionName "Close Presentation" -Action {
                $presentation.Close()
            } -Tags $Tags
        } catch {
            Write-Host "Failed to close presentation: $_"
        }
    }

    if ($ppt -ne $null) {
        try {
            Measure-ExecutionTime -ActionName "Quit PowerPoint Application" -Action {
                $ppt.Quit()
            } -Tags $Tags
        } catch {
            Write-Host "Failed to quit PowerPoint application: $_"
        }
    }

    if ($ppt -ne $null) {
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
        } catch {
            Write-Host "Failed to release COM object."
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    $pptProcess = Get-Process | Where-Object { $_.ProcessName -like "POWERPNT*" }
    if ($pptProcess) {
        Write-Host "Force closing PowerPoint..."
        Stop-Process -Name "POWERPNT" -Force
    }
}

# End overall timer
$overallEnd = Get-Date
$totalElapsed = $overallEnd - $overallStart
Write-Host ("Total execution time: {0:mm} min {0:ss} sec {0:fff} ms" -f $totalElapsed)
