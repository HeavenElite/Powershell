$Computers =  Import-Csv .\ITLab\ITLabData.csv
$Date      =  Get-Date -Format "yyyy.MM.dd"

for ( $i = 0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    $LoopContinue  =  $true

    try {
        $ProArck = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {$Env:PROCESSOR_ARCHITECTURE} -ErrorAction Stop
        $OS      = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {[System.Environment]::OSVersion.VersionString} -ErrorAction Stop
        $OSArck  = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {(wmic os get osarchitecture)[2].Substring(0,2)} -ErrorAction Stop
    }
    catch {
        $Result = @{
            Computer  = $Computers[$i].IP
            Processor = 'PCOffline'
            OSName    = 'PCOffline'
            OSArck    = 'PCOffline'
            Username  = $Computers[$i].Account
            Software  = 'PCOffline'
            SoftArck  = 'PCOffline'
            Version   = 'PCOffline'
        }
        [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Format-Table -AutoSize
        [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport-$Date.csv" -Append -NoTypeInformation
    
        $LoopContinue = $false
    }
    finally {
    
    if ($LoopContinue) {

        if ($ProArck -match '86') {
    
            $Response = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' | Get-ItemProperty | Select-Object -Property DisplayName,DisplayVersion} | Where-Object {$_.Displayname -ne $null}
    
            if ($null -ne $Response) {
    
                [Int32]$Count = ($Response | Measure-Object).Count

                for ($m=0; $m -lt $Count; $m++) {
    
                    $Result = @{
    
                        Computer  = $Computers[$i].IP
                        Processor = $ProArck
                        OSName    = $OS
                        OSArck    = $OSArck
                        Username  = $Computers[$i].Account
                        Software  = $Response[$m].DisplayName
                        SoftArck  = '32-Bit'
                        Version   = $Response[$m].DisplayVersion
                    }
    
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Format-Table -AutoSize
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport-$Date.csv" -Append -NoTypeInformation
                }
            }
    
            else {
    
                $Result = @{
                    Computer  = $Computers[$i].IP
                    Processor = $ProArck
                    OSName    = $OS
                    OSArck    = $OSArck
                    Username  = $Computers[$i].Account
                    Software  = "No32-bitSoftwareFound"
                    SoftArck  = "No32-bitSoftwareFound"
                    Version   = "No32-bitSoftwareFound"
                }
    
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Format-Table -AutoSize
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport-$Date.csv" -Append -NoTypeInformation
            } 
        }
    
        elseif ($ProArck -match '64') {
    
            $Response = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-ChildItem -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\' | Get-ItemProperty | Select-Object -Property DisplayName,DisplayVersion} | Where-Object {$_.Displayname -ne $null}
        
            if ($null -ne $Response) {
    
                [Int32]$Count = ($Response | Measure-Object).Count
        
                for ($m=0; $m -lt $Count; $m++) {
        
                    $Result = @{
        
                        Computer  = $Computers[$i].IP
                        Processor = $ProArck
                        OSName    = $OS
                        OSArck    = $OSArck
                        Username  = $Computers[$i].Account
                        Software  = $Response[$m].DisplayName
                        SoftArck  = '64-Bit'
                        Version   = $Response[$m].DisplayVersion
                    }
        
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Format-Table -AutoSize
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport-$Date.csv" -Append -NoTypeInformation
                }
            }
        
            else {
    
                $Result = @{
                    Computer  = $Computers[$i].IP
                    Processor = $ProArck
                    OSName    = $OS
                    OSArck    = $OSArck
                    Username  = $Computers[$i].Account
                    Software  = "No64-BitSoftwareFound"
                    SoftArck  = "No64-BitSoftwareFound"
                    Version   = "No64-BitSoftwareFound"
                }
    
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Format-Table -AutoSize
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport-$Date.csv" -Append -NoTypeInformation
    
            }
        
            $Response = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' | Get-ItemProperty | Select-Object -Property DisplayName,DisplayVersion} | Where-Object {$_.Displayname -ne $null}
        
            if ($null -ne $Response) {
    
                [Int32]$Count = ($Response | Measure-Object).Count
    
                for ($m=0; $m -lt $Count; $m++) {
    
                    $Result = @{
        
                        Computer  = $Computers[$i].IP
                        Processor = $ProArck
                        OSName    = $OS
                        OSArck    = $OSArck
                        Username  = $Computers[$i].Account
                        Software  = $Response[$m].DisplayName
                        SoftArck  = '32-Bit'
                        Version   = $Response[$m].DisplayVersion
                    }
        
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Format-Table -AutoSize
                    [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport-$Date.csv" -Append -NoTypeInformation
                }
            }
    
            else {
    
                $Result = @{
                    Computer  = $Computers[$i].IP
                    Processor = $ProArck
                    OSName    = $OS
                    OSArck    = $OSArck
                    Username  = $Computers[$i].Account
                    Software  = "No32-BitSoftwareFound"
                    SoftArck  = "No32-BitSoftwareFound"
                    Version   = "No32-BitSoftwareFound"
                }
    
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Format-Table -AutoSize
                [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport-$Date.csv" -Append -NoTypeInformation
            }
        }    
    
        else {
    
            $Result = @{
                Computer  = $Computers[$i].IP
                Processor = $ProArck
                OSName    = $OS
                OSArck    = $OSArck
                Username  = $Computers[$i].Account
                Software  = 'ProcessorInfoError'
                SoftArck  = 'ProcessorInfoError'
                Version   = 'ProcessorInfoError'
            }
    
            [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Format-Table -AutoSize
            [PSCustomObject]$Result | Select-Object -Property Computer,Processor,OSName,OSArck,Username,Software,SoftArck,Version | Export-Csv -Path ".\SoftwareListReport-$Date.csv" -Append -NoTypeInformation
            
                
        }
    }
    }
}

[System.Console]::Beep(1000,1000)