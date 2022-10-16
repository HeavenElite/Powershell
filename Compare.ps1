function Set-Analyze {
    param (
        [CmdLetBinding()]
        [Parameter(Mandatory)]
        [String]$Name

    )
    
    $Result = Compare-Object (Get-Content ".\$($Name)Report-$((Get-Date).AddDays(-3).ToString('yyyy.MM.dd')).csv") (Get-Content ".\$($Name)Report-$((Get-Date).AddDays(-2).ToString('yyyy.MM.dd')).csv")

    for ($i=0; $i -lt $Result.Length/2; $i++) {

        $Report = @{

            TodayResult  = $Result.InputObject[$i]
            YesterResult = $Result.InputObject[$i + $Result.Length/2]
        }

        [PSCustomObject]$Report | Select-Object -Property YesterResult,TodayResult | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property YesterResult,TodayResult | Export-Csv -Path "..\Analyze\$($Name)Analyze-$((Get-Date).ToString('yyyy.MM.dd')).csv" -Append -NoTypeInformation
    }
}

Get-ChildItem . | Where-Object {$_.Name -like '*Report*.csv'} | Move-Item -Destination .\Report\
Set-Location -Path .\Report

Set-Analyze -Name Ping

# Ping
# BackupCheck
# SoftwareEPSCheck
# SoftwareLPECheck
# SoftwareWinSCPCheck



