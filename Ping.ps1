$Servers = Import-Csv -Path .\ITLab\ITLabPing.csv -Header 'Protocal', 'IP', 'Port', 'Site'
$Path    = ".\PingReport-$(Get-Date -Format "yyyy.MM.dd").csv"

if (Test-Path -Path $Path) {

    Remove-Item -Path .\$Path
    Write-Output "Previous file $Path is removed. `n"
}

else {

    Write-Output "Csv file $Path is generated. `n"
}

for ( $i = 0; $i -lt $Servers.Length; $i++)
{   
    $Report = @{
       Site    =  $Servers[$i].Site
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
    
    [PSCustomObject]$Report | Select-Object -Property Site,Server,Service,Port,Result | Format-Table -AutoSize
    [PSCustomObject]$Report | Select-Object -Property Site,Server,Service,Port,Result | Export-Csv -Path $Path -NoTypeInformation -Append
}

Import-Csv -Path $Path | Sort-Object -Property Result,Site | Format-Table -AutoSize

[System.Console]::Beep(1000,1000)