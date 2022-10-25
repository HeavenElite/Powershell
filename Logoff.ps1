$Computers       = Import-Csv -Path ".\ITLab\ITLabData.csv"
$ActiveComputers = Import-Csv -Path ".\UserCheckReport-$(Get-Date -Format 'yyyy.MM.dd').csv" | Where-Object {$_.Status -eq 'Active' -and $_.Session -eq 'Console'}

for ($i=0; $i -lt ($ActiveComputers | Measure-Object).Count; $i++) {

    $IPAddress  =  $ActiveComputers[$i].IPAddress
    $Username   =  $Computers | Where-Object {$_.IP -eq $IPAddress} | Select-Object -ExpandProperty Account
    $Password   =  ConvertTo-SecureString -AsPlainText ($Computers | Where-Object {$_.IP -eq $IPAddress} | Select-Object -ExpandProperty Password) -Force
    $Credential =  New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {logoff.exe $($args[0])} -ArgumentList $ActiveComputers[$i].ID

    $Report = @{

        IPAddress = $IPAddress
        Logoff    = CommandSent
    }

    $Report | Select-Object -Property IPAddress,Logoff | Format-Table -AutoSize
    $Report | Select-Object -Property IPAddress,Logoff | Format-Table -AutoSize | Export-Csv -Path ".\LogoffReport-$(Get-Date -Format 'yyyy.MM.dd').csv"
}