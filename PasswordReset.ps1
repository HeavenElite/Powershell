$Computers = Import-Csv -Path .\ITLab\ITLabDebug.csv
$Count     = ($Computers | Measure-Object).Count

$NewPassword = ConvertTo-SecureString -AsPlainText -Force 'Laurence'

for ($i=0; $i -lt $Count; $i++) {

    $Username    = $Computers[$i].Account
    $OriPassword = ConvertTo-SecureString -AsPlainText -Force $Computers[$i].Password
    $Credential  = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$OriPassword


    Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Set-LocalUser -Name $using:Username -Password $using:NewPassword}
}