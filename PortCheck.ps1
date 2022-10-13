$Servers = Import-Csv -Path .\ITLabService.csv -Header 'Protocal', 'IP', 'Port'
$i = 1
foreach ($Server in $Servers)
{   
    $Report = @{
       No      =  $i
       Server  =  $Server.IP
       Service =  $Server.Protocal
       Port    =  $Server.Port
       Result  =  $null
    }

    if ((Test-NetConnection $Server.IP -Port $Server.Port -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Tcptestsucceeded) {

        $Report.Result = $true
    }

    else {
    
        $Report.Result = $false
    }
    
    [PSCustomObject]$Report | Select-Object -Property No,Server,Service,Port,Result | Format-Table -AutoSize
    [PSCustomObject]$Report | Select-Object -Property No,Server,Service,Port,Result | Export-Csv -Path ".\PortCheckReport-$(Get-Date -Format "yyyy.MM.dd").csv" -NoTypeInformation -Append

    $i++
}
[System.Console]::Beep(1000,1000)