﻿function Get-Storage {
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
       $Result | Select-Object -Property PSComputerName,DeviceID,Size,FreeSpace | Where-Object {$_.FreeSpace -ne '0'} | Export-Csv -Path ".\PracVM-$(Get-Date -Format "dddd.HH.mm").csv" -Append -NoTypeInformation   
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

    $Date = Get-Date -Format 'yyyy.MM.dd'
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
        [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Export-Csv -Path "D:\Desktop\Powershell\RDPReport-$Date.csv" -Append -NoTypeInformation
    }
}
function SoftwareList {

    $LoopContinue = $true

    try {
        $ProArck = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {$Env:PROCESSOR_ARCHITECTURE} -ErrorAction Stop
        $OS      = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {[System.Environment]::OSVersion.VersionString} -ErrorAction Stop
        $OSArck  = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {(wmic os get osarchitecture)[2] -replace ' ',''} -ErrorAction Stop
    }
    catch {
        $Result = @{
            Computer  = $IPAddress
            Processor = 'PCOffline'
            OSName    = 'PCOffline'
            OSArck    = 'PCOffline'
            Software  = 'PCOffline'
            SoftArck  = 'PCOffline'
            Version   = 'PCOffline'
        }
        [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
        [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport.csv" -Append -NoTypeInformation

        $LoopContinue = $false
    }
    finally {

    if ($LoopContinue) {

        if ($ProArck -match '86') {

            $Response = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' | Get-ItemProperty | Select-Object -Property DisplayName,DisplayVersion} | Where-Object {$null -ne $_.Displayname}

            if ($null -ne $Response) {

                [Int32]$Count = ($Response | Measure-Object).Count

                for ($i=0; $i -lt $Count; $i++) {

                    $Result = @{

                        Computer     = $IPAddress
                        Processor    = $ProArck
                        OSName       = $OS
                        OSArck       = $OSArck
                        Software     = $Response[$i].DisplayName
                        SoftArck     = '32-Bit'
                        Version      = $Response[$i].DisplayVersion
                    }

                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport.csv" -Append -NoTypeInformation
                }
            }

            else {

                $Result = @{
                    Computer     = $IPAddress
                    Processor    = $ProArck
                    OSName       = $OS
                    OSArck       = $OSArck
                    Software     = "No32-bitSoftwareFound"
                    SoftArck     = "No32-bitSoftwareFound"
                    Version      = "No32-bitSoftwareFound"
                }

                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport.csv" -Append -NoTypeInformation
            } 
        }

        elseif ($ProArck -match '64') {

            $Response = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {Get-ChildItem -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\' | Get-ItemProperty | Select-Object -Property DisplayName,DisplayVersion} | Where-Object {$null -ne $_.Displayname}
        
            if ($null -ne $Response) {

                [Int32]$Count = ($Response | Measure-Object).Count
        
                for ($i=0; $i -lt $Count; $i++) {
        
                    $Result = @{
        
                        Computer     = $IPAddress
                        Processor    = $ProArck
                        OSName       = $OS
                        OSArck       = $OSArck
                        Software     = $Response[$i].DisplayName
                        SoftArck     = '64-Bit'
                        Version      = $Response[$i].DisplayVersion
                    }
        
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport.csv" -Append -NoTypeInformation
                }
            }
        
            else {

                $Result = @{
                    Computer     = $IPAddress
                    Processor    = $ProArck
                    OSName       = $OS
                    OSArck       = $OSArck
                    Software     = "No64-BitSoftwareFound"
                    SoftArck     = "No64-BitSoftwareFound"
                    Version      = "No64-BitSoftwareFound"
                }

                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport.csv" -Append -NoTypeInformation

            }
        
            $Response = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' | Get-ItemProperty | Select-Object -Property DisplayName,DisplayVersion} | Where-Object {$null -ne $_.Displayname}
        
            if ($null -ne $Response) {

                [Int32]$Count = ($Response | Measure-Object).Count

                for ($i=0; $i -lt $Count; $i++) {

                    $Result = @{
        
                        Computer     = $IPAddress
                        Processor    = $ProArck
                        OSName       = $OS
                        OSArck       = $OSArck
                        Software     = $Response[$i].DisplayName
                        SoftArck     = '32-Bit'
                        Version      = $Response[$i].DisplayVersion
                    }
        
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport.csv" -Append -NoTypeInformation
                }
            }

            else {

                $Result = @{
                    Computer     = $IPAddress
                    Processor    = $ProArck
                    OSName       = $OS
                    OSArck       = $OSArck
                    Software     = "No32-BitSoftwareFound"
                    SoftArck     = "No32-BitSoftwareFound"
                    Version      = "No32-BitSoftwareFound"
                }

                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport.csv" -Append -NoTypeInformation
            }
        }    

        else {

            $Result = @{
                Computer     = $IPAddress
                Processor    = $ProArck
                OSName       = $OS
                OSArck       = $OSArck
                Software     = 'ProcessorInfoError'
                SoftArck     = 'ProcessorInfoError'
                Version      = 'ProcessorInfoError'
            }

            [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Format-Table -AutoSize
            [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport.csv" -Append -NoTypeInformation
        }
    }
    }
}
function UserCheck {

    try {
        $Response = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {qwinsta.exe /server:localhost | Select-String -Pattern '会话','Session','rdp-tcp#','运行中','Active' | ForEach-Object {$_ -replace ' +',' '}} -ErrorAction Stop
    }
    catch {
        $Report = @{

            IPAddress = $IPAddress
            Account = '设备离线'
            Session = '设备离线'
            Status = '设备离线' 
            Suggestion = '检查硬件'
        }

        [PSCustomObject]$Report | Select-Object -Property 'IPAddress','Account','Session','Status','Suggestion' | Format-Table -AutoSize
    }
    finally {

        if (($Response | Measure-Object).Count -eq 1) {
            
            $Report = @{

                IPAddress = $IPAddress
                Account = '空闲'
                Session = '空闲'
                Status = '空闲' 
                Suggestion = '登录远程桌面'
            }
            [PSCustomObject]$Report | Select-Object -Property 'IPAddress','Account','Session','Status','Suggestion' | Format-Table -AutoSize 
        }

        elseif (($Response | Measure-Object).Count -gt 1) {

            for ($i=1; $i -lt $Response.Length; $i++) {
                $Report = @{

                    IPAddress = $IPAddress
                    Account = $Response[$i].Split(' ')[2]
                    Session = $Response[$i].Split(' ')[1]
                    Status = $Response[1].Split(' ')[4]
                    Suggestion = '联系登录用户'
                }
            
            [PSCustomObject]$Report | Select-Object -Property 'IPAddress','Account','Session','Status','Suggestion' | Format-Table -AutoSize
            }
        }

        else {

            $Report = @{

                IPAddress = $IPAddress
                Account = '执行异常'
                Session = '执行异常'
                Status = '执行异常'
                Suggestion = '执行异常'
            }
        
            [PSCustomObject]$Report | Select-Object -Property 'IPAddress','Account','Session','Status','Suggestion' | Format-Table -AutoSize
        }
    }
}

# FunctionList
# Get-Storage
# Show-Backup
# Shutdown
# PSVersion
# AutoStart
# ServiceCheck
# SoftwareCheck -Name
# BackupCheck
# Ping -Port
# RDPRecord
# SoftwareList
# UserCheck

$IPAddress  = '192.168.0.160'
$Username   = Import-Csv -Path .\ITLabData.csv | Where-Object {$_.IP -eq "$IPAddress"} | Select-Object -ExpandProperty Account
$Password   = ConvertTo-SecureString -AsPlainText -Force (Import-Csv -Path .\ITLabData.csv | Where-Object {$_.IP -eq "$IPAddress"} | Select-Object -ExpandProperty Password)
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

RDPRecord