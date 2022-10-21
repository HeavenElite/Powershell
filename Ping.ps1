$Servers = Import-Csv -Path .\ITLab\ITLabPing.csv -Header 'Protocal', 'IP', 'Port'

for ( $i = 0; $i -lt $Servers.Length; $i++)
{   
    $Report = @{
       No      =  $i + 1
       Server  =  $Servers[$i].IP
       Service =  $Servers[$i].Protocal
       Port    =  $Servers[$i].Port
       Result  =  $null
    }

    if ((Test-NetConnection $Servers[$i].IP -Port $Servers[$i].Port -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Tcptestsucceeded) {

        $Report.Result = $true
    }

    else {
    
        $Report.Result = $false
    }
    
    [PSCustomObject]$Report | Select-Object -Property No,Server,Service,Port,Result | Format-Table -AutoSize
    [PSCustomObject]$Report | Select-Object -Property No,Server,Service,Port,Result | Export-Csv -Path ".\PingReport-$(Get-Date -Format "yyyy.MM.dd").csv" -NoTypeInformation -Append
}

[System.Console]::Beep(1000,1000)