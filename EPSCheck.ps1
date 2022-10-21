$Computers = Import-Csv .\ITLabSecure.csv
for ( $i = 0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    try {

        $Result = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -FilePath .\EPSCheckRemoteScript.ps1 -ErrorAction Stop
        $Report = @{
            Computer =  $Computers[$i].IP
            Result   =  $Result
        }
    }
    catch {

        $Report = @{
            Computer =  $Computers[$i].IP
            Result   =  'PCOffline'
        }
    }
    finally {
        
        [PSCustomObject]$Report | Select-Object -Property Computer,Result | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property Computer,Result | Export-Csv -Path ".\SoftwareEPSCheckReport-$(Get-Date -Format "yyyy.MM.dd").csv" -Append -NoTypeInformation
    }
}

[System.Console]::Beep(1000,1000)