$Computers = Import-Csv .\ITLab\ITLabData.csv
for ( $i = 0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {w32tm /config /manualpeerlist:"ntp.ntsc.ac.cn time.windows.com" /syncfromflags:manual /update} -ErrorAction Ignore
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {w32tm /query /source} -ErrorAction Ignore
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {W32tm /resync /force} -ErrorAction Ignore
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {W32tm /register} -ErrorAction Ignore
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Start-Service W32time} -ErrorAction Ignore
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {sc.exe config W32time start= auto} -ErrorAction Ignore
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {(sc.exe qc w32time | Select-String -Pattern 'START_TYPE.*' | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value).Split(':')[1]} -ErrorAction Ignore
    $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-Date -Format "yyyy.MM.dd-HH:mm"} -ErrorAction Ignore
#   $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-Service W32time | Select-Object -ExpandProperty Status} -ErrorAction Ignore

    $Report = @{
        Computer =  $Computers[$i].IP
        Result   =  $Result
    }

    [PSCustomObject]$Report | Select-Object -Property Computer,Result | Format-Table -AutoSize
    [PSCustomObject]$Report | Select-Object -Property Computer,Result | Export-Csv -Path ".\DateCheckReport-$(Get-Date -Format "yyyy.MM.dd").csv" -Append -NoTypeInformation

}

Import-Csv ".\DateCheckReport-$(Get-Date -Format "yyyy.MM.dd").csv" | Sort-Object -Property Result | Format-Table -AutoSize

[System.Console]::Beep(1000,1000)