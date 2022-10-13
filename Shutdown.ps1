$Computers = Import-Csv .\ITLabStop.csv
for ( $i = 0; $i -lt $Computers.Length; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Stop-Computer -Force}
}

[System.Console]::Beep(1000,1000)