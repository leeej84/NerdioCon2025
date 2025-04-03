#RDP Logon Simulator

try {
    try {
        $userDetails = Import-Csv -Path .\Users.csv #User CSV File - Server,Domain,UserName,Password,WaitTime
    } catch {
        Throw "There was an error importing the CSV file of users to logon, check that it exists and is in the correct format"
    } 
    

    #Loop through all the users listed in the CSV and log them onto the relevant servers
    foreach ($user in $userDetails) {
        #Add credentials to the credential store
        Write-Output "$(Get-Date -Format 'yy-MM-dd-HH:mm:ss'): Adding user $($user.UserName) to credential cache"
        Invoke-Expression -Command "cmdkey /generic:TERMSRV/$($user.Server) /user:$($user.Domain)\$($user.UserName) /pass:$($user.Password)" | Out-Null
        Start-Sleep -Seconds 1

        #Start the RDP connection
        Write-Output "$(Get-Date -Format 'yy-MM-dd-HH:mm:ss'): Starting RDP session for user $($user.UserName) on server $($user.Server)"
        Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$($user.Server)" | Out-Null
        
        #Wait time before initiating the next logon
        Write-Output "$(Get-Date -Format 'yy-MM-dd-HH:mm:ss'): Waiting for $($user.WaitTime) seconds before the next logon"
        Start-Sleep -Seconds $($user.WaitTime)
        
        #Remove credentials from the credentials store
        Write-Output "$(Get-Date -Format 'yy-MM-dd-HH:mm:ss'): Removing user $($user.UserName) from credential cache"
        Invoke-Expression -Command "cmdkey /delete:TERMSRV/$($user.Server)" | Out-Null
        Start-Sleep -Seconds 1
    }
} catch {
    Write-Error $_
}
