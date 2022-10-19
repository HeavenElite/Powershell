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
        [String]$Name
    )

    $Result   = Invoke-Command -ComputerName $IPAddress -Credential $Credential -FilePath .\SoftwareCheckRemoteScript.ps1 -ArgumentList $Name,$IPAddress

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
function RDPRecord {

    $Data = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {Get-WinEvent -LogName 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'} | Where-Object {$_.ID -eq '1149' -and $_.TimeCreated -like "*$(Get-Date -Format 'MM/dd/yyyy')*"}

    for ($m=0; $m -lt $Data.Length; $m++) {
    
        $Report = @{
    
            User      = ($Data[$m].Message | Select-String -Pattern '地址:.*').Matches.Value -replace '地址: ',''
            Device    = $IPAddress
            Account   = ($Data[$m].Message | Select-String -Pattern '用户: .*').Matches.Value -replace '用户: ','' -replace "`r",''
            Domain    = ($Data[$m].Message | Select-String -Pattern '域: .*').Matches.Value -replace '域: ','' -replace "`r",''
            LoginTime = $Data[$m].TimeCreated
    
        }
        [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Export-Csv -Path "D:\Desktop\Powershell\RDPReport$(Get-Date -Format 'MM.dd.yyyy').csv" -Append -NoTypeInformation
    }
}

$IPAddress  = '192.168.0.122'
$Username   = 'SHAdmin'
$Password   = ConvertTo-SecureString -AsPlainText -Force 'ShellLPE!23'
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

