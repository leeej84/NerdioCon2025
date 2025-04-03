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

# Start overall timer
$overallStart = Get-Date

# Initialize variables
$word = $null
$doc = $null

try {
    $word = Measure-ExecutionTime -ActionName "Start Word Application" -Action {
        $app = New-Object -ComObject Word.Application
        $app.Visible = $true
        $app.WindowState = 0  # 0 = wdWindowStateNormal

        if ($app.Documents.Count -eq 0) {
            $app.Documents.Add() | Out-Null
        }

        $app.ActiveWindow.Top = 0
        $app.ActiveWindow.Left = 0
        $app.ActiveWindow.Width = 1024
        $app.ActiveWindow.Height = 768

        $app.Activate()
        return $app
    } -Tags $Tags

    if ($null -eq $word) { throw "Failed to start Word application." }

    $doc = Measure-ExecutionTime -ActionName "Create New Document" -Action {
        $word.Documents.Add()
    } -Tags $Tags

    if ($null -eq $doc) { throw "Failed to create new document." }

    Start-Sleep -Seconds 2

    Measure-ExecutionTime -ActionName "Add Text to Document" -Action {
        $doc.Content.Text = "This is an automated test document. Word automation with PowerShell is running."
    } -Tags $Tags

    Measure-ExecutionTime -ActionName "Scroll to End of Document" -Action {
        $doc.Application.Selection.EndKey(6)
    } -Tags $Tags

    Start-Sleep -Seconds 2

    $savePath = "$env:TEMP\TestDocument.docx"
    Measure-ExecutionTime -ActionName "Save Document" -Action {
        $doc.SaveAs2($savePath)
    } -Tags $Tags

    Start-Sleep -Seconds 2
    Write-Host "Automation completed successfully. Document saved at: $savePath"
} catch {
    Write-Host "An error occurred: $_"
} finally {
    if ($doc -ne $null) {
        try {
            Measure-ExecutionTime -ActionName "Close Document" -Action {
                $doc.Close($false)
            } -Tags $Tags
        } catch {
            Write-Host "Failed to close document: $_"
        }
    }

    if ($word -ne $null) {
        try {
            Measure-ExecutionTime -ActionName "Quit Word Application" -Action {
                $word.Quit()
            } -Tags $Tags
        } catch {
            Write-Host "Failed to quit Word application: $_"
        }
    }

    if ($word -ne $null) {
        try {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
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
