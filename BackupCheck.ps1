$Computers = Import-Csv .\ITLabData.csv
for ( $i = 0; $i -lt $Computers.Length; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    try {
        $Result =  Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {wbadmin.exe get versions} -ErrorAction Stop
        $Exist  =  ($Result).Length - 1
        $Length =  ($Result | Select-String '备份时间').Length - 1

        if ( $Exist -gt 8 ) {

            $Report = @{
                Computer =  $Computers[$i].IP
                BackupItem      =   ($Result | Select-String '可以恢复')[$Length]
                OldestVersion   =   ($Result | Select-String '备份时间')[0]
                LatestVersion   =   ($Result | Select-String '备份时间')[$Length]
                VersionNumber   =   $Length + 1
                BackupLocation  =   ($Result | Select-String '备份目标')[$Length]
            
            }
        }
        elseif ( $Exist -eq 8 ) {

            $Report = @{
                Computer =  $Computers[$i].IP
                BackupItem      =   $Result | Select-String '可以恢复'
                OldestVersion   =   'N/A'
                LatestVersion   =   $Result | Select-String '备份时间'
                VersionNumber   =   1
                BackupLocation  =   $Result | Select-String '备份目标'
            
            }
        }
        else {

            $Report = @{
                Computer =  $Computers[$i].IP
                BackupItem      =   'N/A'
                OldestVersion   =   'N/A'
                LatestVersion   =   'N/A'
                VersionNumber   =   0
                BackupLocation  =   'N/A'    
            }
        }
    }
    catch {

        $Report = @{
            Computer =  $Computers[$i].IP
            BackupItem      =   'PCOffline'
            OldestVersion   =   'PCOffline'
            LatestVersion   =   'PCOffline'
            VersionNumber   =   'PCOffline'
            BackupLocation  =   'PCOffline'
       }
    }
    finally {
        [PSCustomObject]$Report | Select-Object -Property Computer,BackupItem,OldestVersion,LatestVersion,VersionNumber,BackupLocation | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property Computer,BackupItem,OldestVersion,LatestVersion,VersionNumber,BackupLocation | Export-Csv -Path ".\BackupCheckReport-$(Get-Date -Format "yyyy.MM.dd").txt" -Append -NoTypeInformation
    }

}

[System.Console]::Beep(1000,1000)