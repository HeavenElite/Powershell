$Computers = Import-Csv -Path .\ITLab\ITLabData.csv | Where-Object {$_.EPS -eq 'EPS'}
$Path      = ".\ConfigEPSCheckReport-$(Get-Date -Format 'yyyy.MM.dd').csv"

for ($i=0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username    = $Computers[$i].Account
    $Password    = ConvertTo-SecureString -AsPlainText -Force $Computers[$i].Password
    $Credential  = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password
    
    $LoopContinue = $true

    try {
        $Response = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -FilePath .\ConfigEPSCheckRemoteScript.ps1 -ArgumentList $Computers[$i].Site -ErrorAction Stop
    }

    catch {
        
        $Report = @{

            Environment = $Computers[$i].Test
            SiteID    = $Computers[$i].Site
            Type      = $Computers[$i].Type
            IPAddress = $Computers[$i].IP
            ForwardServerIP   = "Offline"
            ForwardServerPort = "Offline"
            FuelServerIP = "Offline"
            FuelServerPort = "Offline"
            ConfigSiteID   = "Offline"
        }

        [PSCustomObject]$Report | Select-Object -Property Environment,SiteID,Type,IPAddress,ForwardServerIP,ForwardServerPort,FuelServerIP,FuelServerPort,ConfigSiteID | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property Environment,SiteID,Type,IPAddress,ForwardServerIP,ForwardServerPort,FuelServerIP,FuelServerPort,ConfigSiteID | Export-Csv -Path $Path -Append -NoTypeInformation

        $LoopContinue = $false
    }

    finally {
    if ($LoopContinue) {
        
        $Report = @{

                Environment = $Computers[$i].Test
                SiteID    = $Computers[$i].Site
                Type      = $Computers[$i].Type
                IPAddress = $Computers[$i].IP
                ForwardServerIP   = $Response.ForwardServerIP
                ForwardServerPort = $Response.ForwardServerPort
                FuelServerIP = $Response.FuelServerIP
                FuelServerPort = $Response.FuelServerPort
                ConfigSiteID   = $Response.SiteID
            }

            [PSCustomObject]$Report | Select-Object -Property Environment,SiteID,Type,IPAddress,ForwardServerIP,ForwardServerPort,FuelServerIP,FuelServerPort,ConfigSiteID | Format-Table -AutoSize
            [PSCustomObject]$Report | Select-Object -Property Environment,SiteID,Type,IPAddress,ForwardServerIP,ForwardServerPort,FuelServerIP,FuelServerPort,ConfigSiteID | Export-Csv -Path $Path -Append -NoTypeInformation
    }
    }
}

Import-Csv -Path $Path | Sort-Object -Property Environment,SiteID,Type | Format-Table -AutoSize

[System.Console]::Beep(1000,1000)