$Computers = Import-Csv -Path .\ITLab\ITLabDebug.csv
$Count     = ($Computers | Measure-Object).Count

$Path        = "PasswordResetReport-$(Get-Date -Format "yyyy.MM.dd").csv"
$NewPassword = ConvertTo-SecureString -AsPlainText 'Laurence' -Force

for ($i=0; $i -lt $Count; $i++) {

    $Username    = $Computers[$i].Account
    $OriPassword = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential  = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$OriPassword

    $LoopContinue = $true

    try {
        Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Set-LocalUser -Name $using:Username -Password $using:NewPassword} -ErrorAction Stop
    }

    catch {
        $Report = @{
            Environment   = $Computers[$i].Test
            SiteID        = $Computers[$i].Site
            DeviceType    = $Computers[$i].Type
            IPAddress     = $Computers[$i].IP
            Operation     = 'ComputerOffline'
        }
        [PSCustomObject]$Report | Select-Object -Property Environment,SiteID,DeviceType,IPAddress,Operation | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property Environment,SiteID,DeviceType,IPAddress,Operation | Export-Csv -Path $Path -Append -NoTypeInformation

        $LoopContinue = $false
    }

    finally {
    if ($LoopContinue) {

        $Report = @{
            Environment   = $Computers[$i].Test
            SiteID        = $Computers[$i].Site
            DeviceType    = $Computers[$i].Type
            IPAddress     = $Computers[$i].IP
            Operation     = 'ResetCommandSent'
        }
        [PSCustomObject]$Report | Select-Object -Property Environment,SiteID,DeviceType,IPAddress,Operation | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property Environment,SiteID,DeviceType,IPAddress,Operation | Export-Csv -Path $Path -Append -NoTypeInformation
    }
    }
}

Import-Csv -Path $Path | Sort-Object -Property Operation,Environment,SiteID,DeviceType | Format-Table -AutoSize

[System.Console]::Beep(1000,1000)