$Computers = Import-Csv .\ITLabData.csv
for ($i=0; $i -lt $Computers.Length; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    $Data = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-WinEvent -LogName 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'} | Where-Object {$_.ID -eq '1149' -and $_.TimeCreated -like "*$(Get-Date -Format 'MM/dd/yyyy')*"}

    for ($m=0; $m -lt $Data.Length; $m++) {
    
        $Report = @{
    
            User      = ($Data[$m].Message | Select-String -Pattern '地址:.*').Matches.Value -replace '地址: ',''
            Device    = $Computers[$i].IP
            Account   = ($Data[$m].Message | Select-String -Pattern '用户: .*').Matches.Value -replace '用户: ','' -replace "`r",''
            Domain    = ($Data[$m].Message | Select-String -Pattern '域: .*').Matches.Value -replace '域: ','' -replace "`r",''
            LoginTime = $Data[$m].TimeCreated
    
        }
        [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Export-Csv -Path "D:\Desktop\Powershell\RDPReport$(Get-Date -Format 'yyyy.MM.dd').csv" -Append -NoTypeInformation

    }
}