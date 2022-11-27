$Computers = Import-Csv .\ITLab\ITLabStop.csv
for ( $i = 0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Stop-Computer -Force}

    $Report = @{

        IPAddress = $IPAddress
        Shutdown  = 'CommandSent'
    }

    $Report | Select-Object -Property IPAddress,Shutdown | Format-Table -AutoSize
    $Report | Select-Object -Property IPAddress,Shutdown | Export-Csv -Path ".\ShutdownReport-$(Get-Date -Format 'yyyy.MM.dd').csv" -Append -NoTypeInformation
}

[System.Console]::Beep(1000,1000)