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
$OriPath = ".\Analyze\$($Name)Analyze-$((Get-Date).ToString('yyyy.MM.dd')).csv"
$ForPath = ".\Analyze\$($Name)Analyze-$((Get-Date).ToString('yyyy.MM.dd')).Format.csv"

if ($Name -match 'BackupCheck') {
   
    (Get-Content $OriPath) -replace '^"""|"""$','"' -replace '""','"'  -replace '","',"`t" -replace '可以恢复: 卷， 文件， 应用程序， 裸机恢复， 系统状态|备份时间: |新加卷|本地磁盘','' -replace '备份目标: 1394/USB 磁盘，标签为 ','1394/USB' -replace '备份目标: 固定磁盘，标签为 ','LocalDisk' | Set-Content $OriPath
    (Get-Content $OriPath) -replace "`t",',' -replace ',",',',"",' | Set-Content $OriPath

    $Data = Import-Csv $OriPath -Header PreComputer,PreBackupItem,PreOldestVersion,PreLatestVersion,PreVersionNumber,PreBackupLocation,CurComputer,CurBackupItem,CurOldestVersion,CurLatestVersion,CurVersionNumber,CurBackupLocation 

    for ($i=1; $i -lt $Data.Length; $i++) {
        $Report = @{
            PreComputer         =  $Data.PreComputer[$i]
            PreBackupItem       =  $Data.PreBackupItem[$i]
            PreOldestVersion    =  $Data.PreOldestVersion[$i]
            PreLatestVersion    =  $Data.PreLatestVersion[$i]
            PreVersionNumber    =  $Data.PreVersionNumber[$i]
            PreBackupLocation   =  $Data.PreBackupLocation[$i]
            CurComputer         =  $Data.CurComputer[$i]
            CurBackupItem       =  $Data.CurBackupItem[$i]
            CurOldestVersion    =  $Data.CurOldestVersion[$i]
            CurLatestVersion    =  $Data.CurLatestVersion[$i]
            CurVersionNumber    =  $Data.CurVersionNumber[$i]
            CurBackupLocation   =  $Data.CurBackupLocation[$i]
        }

        [PSCustomObject]$Report | Select-Object -Property PreComputer,PreBackupItem,PreOldestVersion,PreLatestVersion,PreVersionNumber,PreBackupLocation,CurComputer,CurBackupItem,CurOldestVersion,CurLatestVersion,CurVersionNumber,CurBackupLocation | Export-Csv $ForPath -Append -NoTypeInformation
        
    }

    (Get-Content $ForPath) -replace ',',"`t" | Set-Content $ForPath
}

elseif ($Name -match 'Ping') {
    
    (Get-Content $OriPath) -replace '^"""|"""$','"' -replace '","',"`t" -replace '""','"' | Set-Content $OriPath
    (Get-Content $OriPath) -replace "`t",',' | Set-Content $OriPath

    $Data = Import-Csv $OriPath -Header PreNo,PreServer,PreService,PrePort,PreResult,CurNo,CurServer,CurService,CurPort,CurResult

    for ($i=1; $i -lt $Data.Length; $i++) {

        $Report = @{
            PreNo       =  $Data.PreNo[$i]
            PreServer   =  $Data.PreServer[$i]
            PreService  =  $Data.PreService[$i]
            PrePort     =  $Data.PrePort[$i]
            PreResult   =  $Data.PreResult[$i]
            CurNo       =  $Data.CurNo[$i]
            CurServer   =  $Data.CurServer[$i]
            CurService  =  $Data.CurService[$i]
            CurPort     =  $Data.CurPort[$i]
            CurResult   =  $Data.CurResult[$i]
        }

        [PSCustomObject]$Report | Select-Object -Property PreNo,PreServer,PreService,PrePort,PreResult,CurNo,CurServer,CurService,CurPort,CurResult | Export-Csv $ForPath -Append -NoTypeInformation
    }
    (Get-Content $ForPath) -replace ',',"`t" | Set-Content $ForPath

}

elseif ($Name -match 'SoftwareEPSCheck') {

    (Get-Content $OriPath) -replace '^"""|"""$','"' -replace '","',"`t" -replace '""','"' | Set-Content $OriPath
    (Get-Content $OriPath) -replace "`t",',' | Set-Content $OriPath

    $Data = Import-Csv $OriPath -Header PreComputer,PreResult,CurComputer,CurResult

    for ($i=1; $i -lt $Data.Length; $i++) {
        $Report = @{
            PreComputer = $Data.PreComputer[$i]
            PreResult   = $Data.PreResult[$i]
            CurComputer = $Data.CurComputer[$i]
            CurResult   = $Data.CurResult[$i]
        }

        [PSCustomObject]$Report | Select-Object -Property PreComputer,PreResult,CurComputer,CurResult | Export-Csv $ForPath -Append -NoTypeInformation
    }
    (Get-Content $ForPath) -replace ',',"`t" | Set-Content $ForPath

}

elseif ($Name -match 'SoftwareLPECheck' -or 'SoftwareWinSCPCheck') {

    (Get-Content $OriPath) -replace '^"""|"""$','"' -replace '          ','' -replace '","',"`t" -replace '",,"',"`t`t" -replace '""','"' -replace ',"',"`t" | Set-Content $OriPath
    (Get-Content $OriPath) -replace "`t",',' | Set-Content $OriPath

    $Data = Import-Csv $OriPath -Header PreComputer,PreProcessor,PreOSName,PreOSArck,PreSoftware,PreSoftArck,PreVersion,CurComputer,CurProcessor,CurOSName,CurOSArck,CurSoftware,CurSoftArck,CurVersion

    for ($i=1; $i -lt $Data.Length; $i++) {

        $Report = @{

            PreComputer   =  $Data.PreComputer[$i]
            PreProcessor  =  $Data.PreProcessor[$i]
            PreOSName     =  $Data.PreOSName[$i]
            PreOSArck     =  $Data.PreOSArck[$i]
            PreSoftware   =  $Data.PreSoftware[$i]
            PreSoftArck   =  $Data.PreSoftArck[$i]
            PreVersion    =  $Data.PreVersion[$i]
            CurComputer   =  $Data.CurComputer[$i]
            CurProcessor  =  $Data.CurProcessor[$i]
            CurOSName     =  $Data.CurOSName[$i]
            CurOSArck     =  $Data.CurOSArck[$i]
            CurSoftware   =  $Data.CurSoftware[$i]
            CurSoftArck   =  $Data.CurSoftArck[$i]
            CurVersion    =  $Data.CurVersion[$i]
        }

        [PSCustomObject]$Report | Select-Object -Property PreComputer,PreProcessor,PreOSName,PreOSArck,PreSoftware,PreSoftArck,PreVersion,CurComputer,CurProcessor,CurOSName,CurOSArck,CurSoftware,CurSoftArck,CurVersion | Export-Csv $ForPath -Append -NoTypeInformation
    }

    (Get-Content $ForPath) -replace ',',"`t" | Set-Content $ForPath
}

else {
    
    Write-Output 'There must be something Wrong!'
}

Remove-Item -Path $OriPath