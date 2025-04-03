$downloadUrl = "https://github.com/influxdata/telegraf/releases"
$outputPath = "C:\Temp"
$installPath = "C:\Program Files\Telegraf"

#Check for the Influx director first before running or doing anything
if (!(Test-Path $installPath)) {
    try {
        #Set TLS1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

		New-Item $outputPath -ItemType Directory -Force | Out-Null
        $response = Invoke-WebRequest -UseBasicParsing -Uri "https://github.com/influxdata/telegraf/releases"
        $downloadLink = $($response.Links | Where-Object {$_ -match "_windows_amd64"} | Select-Object -First 1).href
        $fileName = $downloadlink.Split("/")[$($downloadLink.split("/").count-1)]
        Invoke-WebRequest -Uri $downloadLink -OutFile "$outputPath\$fileName"
    
        if (Test-Path "$outputPath\$fileName") {
            #Expand the telegraf archive
            Expand-Archive -Path "$outputPath\$fileName" -DestinationPath $installPath -Force
            $exeFile = Get-ChildItem $installPath -Recurse | Where {$_.Name -match "telegraf.exe"}
            $installPath = $exeFile.DirectoryName

            #Setup the telegraf Service
            & $exeFile.FullName --service install
			
			#Download the telegraf config File
			Invoke-WebRequest -UseBasicParsing -Uri "https://raw.githubusercontent.com/leeej84/NerdioCon2025/refs/heads/main/Telegraf/Telegraf.conf" -OutFile "C:\Program Files\Telegraf\telegraf.conf"
			
            #Start the service
            Get-Service telegraf | Start-Service

        } else {
            Throw "Could not download telegraf agent"
        }
    } catch {
        Write-Error $_
        $_ | Out-File "$outputPath\Telegraf_Error.log"
    }
}