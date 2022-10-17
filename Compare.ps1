param (
    [CmdLetBinding()]
    [Parameter(Mandatory)]
    [String]$Name,
    [Parameter(Mandatory)]
    [Int32]$Previous,
    [Parameter(Mandatory)]
    [Int32]$Current
)

try {
    Get-ChildItem . | Where-Object {$_.Name -like '*Report*.csv'} | Move-Item -Destination .\Report\
}
catch{
    Write-Output "There are no reports required to be removed. `n"
}
finally {
    Write-Output "Reports are now removed. `n"
}


Set-Location -Path .\Report

Write-Output ''
Write-Output "The Previous File: $($Name)Report-$((Get-Date).AddDays($Previous).ToString('yyyy.MM.dd')).csv"
Write-Output "The Current File: $($Name)Report-$((Get-Date).AddDays($Current).ToString('yyyy.MM.dd')).csv `n"

$PreData = Get-Content ".\$($Name)Report-$((Get-Date).AddDays($Previous).ToString('yyyy.MM.dd')).csv"
$CurData = Get-Content ".\$($Name)Report-$((Get-Date).AddDays($Current).ToString('yyyy.MM.dd')).csv"

if ($PreData.Length -ne $CurData.Length) {

    Write-Output "Your CSVs are not well aligned and sorted! `n"
    Set-Location -Path ..
    Exit
}

$Result = Compare-Object $PreData $CurData

if ( $null -eq $Result) {

    Write-Output "There is no change in the file today! `n"
    Set-Location -Path ..
    Exit
}

for ($i=0; $i -lt $Result.Length/2; $i++) {

    $Report = @{
        CurrentResult  = $Result.InputObject[$i]
        PreviousResult = $Result.InputObject[$i + $Result.Length/2]
    }

    [PSCustomObject]$Report | Select-Object -Property PreviousResult,CurrentResult | Export-Csv -Path "..\Analyze\$($Name)Analyze-$((Get-Date).ToString('yyyy.MM.dd')).csv" -Append -NoTypeInformation
}

Set-Location -Path ..
$OriPath = ".\Analyze\$($Name)Analyze-$((Get-Date).ToString('yyyy.MM.dd')).csv"
$ForPath = ".\Analyze\$($Name)Analyze-$((Get-Date).ToString('yyyy.MM.dd')).Format.csv"

if ($Name -match 'BackupCheck') {
    
    (Get-Content $OriPath) -replace '^"""|"""$','"' -replace '""','"' -replace '可以恢复: 卷， 文件， 应用程序， 裸机恢复， 系统状态','Barebone' -replace '备份时间: |新加卷|本地磁盘','' -replace '备份目标: 1394/USB 磁盘，标签为 ','1394/USB' -replace '备份目标: 固定磁盘，标签为 ','LocalDisk' -replace '""','"' | Set-Content $OriPath

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
        
        [PSCustomObject]$Report | Select-Object -Property PreComputer,CurComputer,PreBackupItem,CurBackupItem,PreOldestVersion,CurOldestVersion,PreLatestVersion,CurLatestVersion,PreVersionNumber,CurVersionNumber,PreBackupLocation,CurBackupLocation | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property PreComputer,CurComputer,PreBackupItem,CurBackupItem,PreOldestVersion,CurOldestVersion,PreLatestVersion,CurLatestVersion,PreVersionNumber,CurVersionNumber,PreBackupLocation,CurBackupLocation | Export-Csv $ForPath -Append -NoTypeInformation
        
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

        [PSCustomObject]$Report | Select-Object -Property PreNo,CurNo,PreServer,CurServer,PreService,CurService,PrePort,CurPort,PreResult,CurResult | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property PreNo,CurNo,PreServer,CurServer,PreService,CurService,PrePort,CurPort,PreResult,CurResult | Export-Csv $ForPath -Append -NoTypeInformation
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

        [PSCustomObject]$Report | Select-Object -Property PreComputer,CurComputer,PreResult,CurResult | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property PreComputer,CurComputer,PreResult,CurResult | Export-Csv $ForPath -Append -NoTypeInformation
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
        [PSCustomObject]$Report | Select-Object -Property PreComputer,CurComputer,PreProcessor,CurProcessor,PreOSName,CurOSName,PreOSArck,CurOSArck,PreSoftware,CurSoftware,PreSoftArck,CurSoftArck,PreVersion,CurVersion | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property PreComputer,CurComputer,PreProcessor,CurProcessor,PreOSName,CurOSName,PreOSArck,CurOSArck,PreSoftware,CurSoftware,PreSoftArck,CurSoftArck,PreVersion,CurVersion | Export-Csv $ForPath -Append -NoTypeInformation
    }

    (Get-Content $ForPath) -replace ',',"`t" | Set-Content $ForPath
}

else {
    
    Write-Output 'There must be something Wrong!'
}

Write-Output "Analyze Report is generaetd here: $ForPath"
Write-Output ''
Remove-Item -Path $OriPath