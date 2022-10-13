$Computers = Import-Csv .\ITLabData.csv
for ( $i = 0; $i -lt $Computers.Length; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {sc.exe config WinRM start= auto}
    $Report = @{
        Computer = $Computers[$i].IP
        Result   = $Result
    }
    
    [PSCustomObject]$Report | Select-Object -Property Computer,Result | Format-Table -AutoSize
    [PSCustomObject]$Report | Select-Object -Property Computer,Result | Export-Csv -Path ".\ServiceConfigure-$(Get-Date -Format "yyyy.MM.dd").csv" -Append -NoTypeInformation
    }

[System.Console]::Beep(1000,1000)