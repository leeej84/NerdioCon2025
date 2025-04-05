#Define parameters
param (
    [hashtable]$Tags = @{},
    $influxServer
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

    $influxURL = "http://$($influxServer):8086/api/v2/write?org=Performance&bucket=Performance&precision=ms"
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

# Import user32.dll functions to manipulate window visibility and position
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

# Initialize variables
$excel = $null
$workbook = $null

try {
    $excel = Measure-ExecutionTime -ActionName "Start Excel Application" -Action {
        $app = New-Object -ComObject Excel.Application
        $app.Visible = $true
    
        # Force open a workbook to initialize a window
        if ($app.Workbooks.Count -eq 0) {
            $null = $app.Workbooks.Add()
        }
    
        # Use WScript to bring to foreground
        $shell = New-Object -ComObject WScript.Shell
        $shell.AppActivate($app.Caption) | Out-Null
    
        return $app
    } -Tags $Tags

    Start-Sleep -Seconds 1
    $hWnd = [WinAPI]::FindWindow("XLMAIN", $excel.Caption)
    if ($hWnd -ne [IntPtr]::Zero) {
        [WinAPI]::MoveWindow($hWnd, 100, 100, 1024, 768, $true) | Out-Null
    } else {
        Write-Host "Failed to find Excel window handle."
    }


    if ($null -eq $excel) { throw "Failed to start Excel application." }

    $excel.DisplayAlerts = $false

    $workbook = Measure-ExecutionTime -ActionName "Create Workbook" -Action {
        $excel.Workbooks.Add()
    } -Tags $Tags

    if ($null -eq $workbook) { throw "Failed to create a new workbook." }

    Start-Sleep -Seconds 1
    $hWnd = [WinAPI]::FindWindow("XLMAIN", $excel.Caption)
    if ($hWnd -ne [IntPtr]::Zero) {
        [WinAPI]::MoveWindow($hWnd, 100, 100, 1024, 768, $true) | Out-Null
    } else {
        Write-Host "Failed to find Excel window handle."
    }

    Measure-ExecutionTime -ActionName "Ensure Excel is Active" -Action {
        $shell = New-Object -ComObject WScript.Shell
        $shell.AppActivate($excel.Caption) | Out-Null
    } -Tags $Tags

    Measure-ExecutionTime -ActionName "Add Data to Cell" -Action {
        $excel.ActiveSheet.Cells(1,1).Value = "Automated test entry"
    } -Tags $Tags

    Start-Sleep -Seconds 2

    $savePath = "$env:TEMP\TestWorkbook.xlsx"
    if ($workbook -ne $null) {
        Measure-ExecutionTime -ActionName "Save Workbook" -Action {
            $workbook.SaveAs($savePath, 51)
        } -Tags $Tags
    } else {
        Write-Host "Skipping Save: Workbook object is null!"
    }

    Start-Sleep -Seconds 2
    Write-Host "Automation completed successfully. Workbook saved at: $savePath"
} catch {
    Write-Host "An error occurred: $_"
} finally {
    if ($workbook -ne $null) {
        try {
            Measure-ExecutionTime -ActionName "Close Workbook" -Action {
                $workbook.Close($false)
            } -Tags $Tags
        } catch {
            Write-Host "Failed to close workbook: $_"
        }
    }

    if ($excel -ne $null) {
        try {
            Measure-ExecutionTime -ActionName "Quit Excel Application" -Action {
                $excel.Quit()
            } -Tags $Tags
        } catch {
            Write-Host "Failed to quit Excel application: $_"
        }
    }

    if ($excel -ne $null) {
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch {
            Write-Host "Failed to release COM object."
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# End overall timer
$overallEnd = Get-Date
$totalElapsed = $overallEnd - $overallStart
Write-Host ("Total execution time: {0:mm} min {0:ss} sec {0:fff} ms" -f $totalElapsed)
