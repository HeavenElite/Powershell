function Get-Storage {
    [CmdletBinding()]
    $Result = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {Get-PSDrive -PSProvider FileSystem}

    foreach ($Item in $Result) {   
        $Report = @{
            Computer = $Item.PSComputerName
            PartNum  = $Item.Name
            UsedSize = [Math]::Round(($Item.Used / 1GB), 0)
            FreeSize = [Math]::Round(($Item.Free / 1GB), 0)
            PartSize = [Math]::Round((($Item.Used + $Item.Free) / 1GB), 0)
        }
        $Report | Where-Object {$_.FreeSize -ne '0'} | Select-Object -Property Computer,PartNum,PartSize,UsedSize,FreeSize | Format-Table -AutoSize
    #   $Result | Select-Object -Property PSComputerName,DeviceID,Size,FreeSpace | Where-Object {$_.FreeSpace -ne '0'} | Export-Csv -Path ".\PracVM-$(Get-Date -Format "dddd.HH.mm").csv" -Append -NoTypeInformation   
    }
}
function Show-Backup {
    [CmdletBinding()]
    $Result = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {wbadmin.exe get versions}
    $Result
}
function Shutdown {
    [CmdletBinding()]
    $Result = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {Stop-Computer -Force}
    $Result
}
function PSVersion {
    [CmdletBinding()]
    $Result = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {$PSVersionTable}
    $Result
}
function AutoStart {
    [CmdletBinding()]
    $Result = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {sc.exe config WinRM start= auto}
    $Report = @{
        Computer = $IPAddress
        Result   = $Result
    }
    [PSCustomObject]$Report | Select-Object -Property Computer,Result | Format-Table -AutoSize
}
function ServiceCheck {
    [CmdletBinding()]
    $Result = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {sc.exe qc winrm | Select-String "START_TYPE"}
    $Report = @{
        Computer = $IPAddress
        Result   = $Result
    }
    [PSCustomObject]$Report | Select-Object -Property Computer,Result | Format-Table -AutoSize
}
function SoftwareCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]$Software
    )

    $Result   = Invoke-Command -ComputerName $IPAddress -Credential $Credential -FilePath .\SoftwareCheckRemoteScript.ps1 -ArgumentList $Software,$IPAddress

    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
}
function BackupCheck {
    [CmdletBinding()]
    $Result =  Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {wbadmin.exe get versions}
    $Exist  =  ($Result).Length - 1
    $Length =  ($Result | Select-String '备份时间').Length - 1

    if ( $Exist -gt 8 ) {
        $Report = @{
            Computer        =   $IPAddress
            BackupItem      =   ($Result | Select-String '可以恢复')[$Length]
            OldestVersion   =   ($Result | Select-String '备份时间')[0]
            LatestVersion   =   ($Result | Select-String '备份时间')[$Length]
            VersionNumber   =   $Length + 1
            BackupLocation  =   ($Result | Select-String '备份目标')[$Length]
        }
    }
    elseif ( $Exist -eq 8 ) {
        $Report = @{
            Computer        =   $IPAddress
            BackupItem      =   $Result | Select-String '可以恢复'
            OldestVersion   =   'N/A'
            LatestVersion   =   $Result | Select-String '备份时间'
            VersionNumber   =   1
            BackupLocation  =   $Result | Select-String '备份目标'
        }
    }
    else {
        $Report = @{
            Computer        =   $IPAddress
            BackupItem      =   'N/A'
            OldestVersion   =   'N/A'
            LatestVersion   =   'N/A'
            VersionNumber   =   0
            BackupLocation  =   'N/A'
        }
    }
    [PSCustomObject]$Report | Select-Object -Property Computer,BackupItem,OldestVersion,LatestVersion,VersionNumber,BackupLocation | Format-Table -AutoSize
}
function Ping {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]$Port
    )
    Test-NetConnection -ComputerName $IPAddress -Port $Port
}

$IPAddress  = ''
$Username   = ''
$Password   = ConvertTo-SecureString -AsPlainText -Force ''
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password