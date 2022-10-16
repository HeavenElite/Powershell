param (
    [CmdLetBinding()]
    [Parameter(Mandatory)]
    [String]$Name,
    [Parameter(Mandatory)]
    [Int32]$Previous,
    [Parameter(Mandatory)]
    [Int32]$Current
)

Get-ChildItem . | Where-Object {$_.Name -like '*Report*.csv'} | Move-Item -Destination .\Report\
Set-Location -Path .\Report

Write-Output "The Previous File: $($Name)Report-$((Get-Date).AddDays($Previous).ToString('yyyy.MM.dd')).csv"
Write-Output "The Current File: $($Name)Report-$((Get-Date).AddDays($Current).ToString('yyyy.MM.dd')).csv"

$Result = Compare-Object (Get-Content ".\$($Name)Report-$((Get-Date).AddDays($Previous).ToString('yyyy.MM.dd')).csv") (Get-Content ".\$($Name)Report-$((Get-Date).AddDays($Current).ToString('yyyy.MM.dd')).csv")

for ($i=0; $i -lt $Result.Length/2; $i++) {
    $Report = @{
        CurrentResult  = $Result.InputObject[$i]
        PreviousResult = $Result.InputObject[$i + $Result.Length/2]
    }
    [PSCustomObject]$Report | Select-Object -Property PreviousResult,CurrentResult | Format-Table -AutoSize
    [PSCustomObject]$Report | Select-Object -Property PreviousResult,CurrentResult | Export-Csv -Path "..\Analyze\$($Name)Analyze-$((Get-Date).ToString('yyyy.MM.dd')).csv" -Append -NoTypeInformation
}

Set-Location -Path ..

# Ping
# BackupCheck
# SoftwareEPSCheck
# SoftwareLPECheck
# SoftwareWinSCPCheck



