$Computers = Import-Csv .\ITLabData.csv
for ( $i = 0; $i -lt $Computers.Length; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-PSDrive -PSProvider FileSystem} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    

    for ( $m = 0; $m -lt $Result.Length; $m++ ) {

        $Report = @{

            Computer  =  $Result[$m].PSComputerName
            PartNum   =  $Result[$m].Name
            UsedSize  =  [Math]::Round(($Result[$m].Used / 1GB),0)
            FreeSize  =  [Math]::Round(($Result[$m].Free / 1GB),0)
            PartSize  =  [Math]::Round((($Result[$m].Used + $Result[$m].Free) / 1GB),0)

        }

        [PSCustomObject]$Report | Select-Object -Property Computer,PartNum,PartSize,UsedSize,FreeSize | Where-Object {$_.FreeSize -ne '0'} | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property Computer,PartNum,PartSize,UsedSize,FreeSize | Where-Object {$_.FreeSize -ne '0'} | Export-Csv -Path ".\DiskInfo-$(Get-Date -Format "yyyy.MM.dd").csv" -Append -NoTypeInformation
    }
}
[System.Console]::Beep(1000,1000)