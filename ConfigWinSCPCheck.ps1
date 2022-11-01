$Computers = Import-Csv -Path .\ITLab\ITLabData.csv | Where-Object {$_.WinSCP -eq 'WinSCP'}
$Path      = ".\ConfigWinSCPCheckReport-$(Get-Date -Format 'yyyy.MM.dd').csv"

for ($i=0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username    = $Computers[$i].Account
    $Password    = ConvertTo-SecureString -AsPlainText -Force $Computers[$i].Password
    $Credential  = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    $LoopContinue = $true
    
    try {

        $Response = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-Content -Path "C:\WinSCP\Configuration_end.ini"} -ErrorAction Stop
        $Result   = ($Response | Select-String -Pattern '@\d+\.\d+\.\d+\.\d+').Matches.Value.Replace('@','')
    }

    catch {

        $Report = @{
    
            IPAddress   = $Computers[$i].IP
            Environment = $Computers[$i].Test
            SiteID      = $Computers[$i].Site
            Type        = $Computers[$i].Type
            SFTPServer  = "ConfigError Or Offline"
        }
    
        [PSCustomObject]$Report | Select-Object -Property IPAddress,Environment,SiteID,Type,SFTPServer | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property IPAddress,Environment,SiteID,Type,SFTPServer | Export-Csv -Path $Path -Append -NoTypeInformation

        $LoopContinue = $false
    }

    finally {

    if ($LoopContinue) {

        $Report = @{
    
            IPAddress   = $Computers[$i].IP
            Environment = $Computers[$i].Test
            SiteID      = $Computers[$i].Site
            Type        = $Computers[$i].Type
            SFTPServer  = $Result
        }
    
        [PSCustomObject]$Report | Select-Object -Property IPAddress,Environment,SiteID,Type,SFTPServer | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property IPAddress,Environment,SiteID,Type,SFTPServer | Export-Csv -Path $Path -Append -NoTypeInformation
    }
    }
}

Import-Csv -Path $Path | Sort-Object -Property Environment,SiteID | Format-Table -AutoSize

[System.Console]::Beep(1000,1000)