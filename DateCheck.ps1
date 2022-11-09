$Computers = Import-Csv .\ITLab\ITLabData.csv
for ( $i = 0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {w32tm /config /manualpeerlist:"ntp.ntsc.ac.cn time.windows.com" /syncfromflags:manual /update}
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {w32tm /query /source}
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {W32tm /resync /force}
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Start-Service W32time}
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {sc.exe config W32time start= auto}
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {(sc.exe qc w32time | Select-String -Pattern 'START_TYPE.*' | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value).Split(':')[1]}
    $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-Date -Format "yyyy.MM.dd-HH:mm"}
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-Service W32time | Select-Object -ExpandProperty Status}

    $Report = @{
        Computer =  $Computers[$i].IP
        Result   =  $Result
    }

    [PSCustomObject]$Report | Select-Object -Property Computer,Result | Format-Table -AutoSize
    [PSCustomObject]$Report | Select-Object -Property Computer,Result | Export-Csv -Path ".\DateCheckReport-$(Get-Date -Format "yyyy.MM.dd").csv" -Append -NoTypeInformation

}

[System.Console]::Beep(1000,1000)