param(
    [Parameter(Mandatory)]
    [string]$Name
)

$Computers = Import-Csv .\ITLab\ITLabSecure.csv
for ( $i = 0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    try {
        $Result = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -FilePath .\SoftwareCheckRemoteScript.ps1 -ArgumentList $Name,$Computers[$i].IP -ErrorAction Stop
    }
    catch {
        $Result = @{
            Computer  = $Computers[$i].IP
            Processor = 'PCOffline'
            OSName    = 'PCOffline'
            OSArck    = 'PCOffline'
            Software  = 'PCOffline'
            SoftArck  = 'PCOffline'
            Version   = 'PCOffline'
        }
    }
    finally {
        [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
        [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\Software$($Name)CheckReport-$(Get-Date -Format "yyyy.MM.dd").csv" -Append -NoTypeInformation
    }
}

[System.Console]::Beep(1000,1000)