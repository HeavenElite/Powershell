$Computers = Import-Csv -Path .\ITLab\ITLabData.csv | Where-Object {$_.EPS -eq 'EPS'}
$Path      = ".\ConfigEPSCheckReport-$(Get-Date -Format 'yyyy.MM.dd').csv"

for ($i=0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username    = $Computers[$i].Account
    $Password    = ConvertTo-SecureString -AsPlainText -Force $Computers[$i].Password
    $Credential  = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password
    
    $Response = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -FilePath .\ConfigEPSCheckRemoteScript.ps1 -ArgumentList $Computers[$i].IP,$Computers[$i].Site
    $Response | Select-Object -Property IPAddress,ForwardServerIP,FuelServerIP,SiteID,RPOSPort | Format-Table -AutoSize
    $Response | Select-Object -Property IPAddress,ForwardServerIP,FuelServerIP,SiteID,RPOSPort | Export-Csv -Path $Path -Append -NoTypeInformation
}

#Import-Csv -Path $Path | Sort-Object -Property IPAddress,ForwardServerIP,FuelServerIP,SiteID,RPOSPort | Format-Table -AutoSize

[System.Console]::Beep(1000,1000)