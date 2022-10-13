$Software  = $($args[0])
$IPAddress = $($args[1])

if ($Env:PROCESSOR_ARCHITECTURE -match '86') {

    $Response = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' | Get-ItemProperty | Select-Object -Property DisplayName,DisplayVersion | Where-Object {$_.Displayname -like "*$Software*"}

    $Result = @{
        Computer     = $IPAddress
        Processor    = $Env:PROCESSOR_ARCHITECTURE
        OSName       = [System.Environment]::OSVersion.VersionString
        OSArck       = (wmic os get osarchitecture)[2]
        Software     = $Response.DisplayName
        SoftArck     = '32-Bit'
        Version      = $Response.DisplayVersion
    }

}
elseif ($Env:PROCESSOR_ARCHITECTURE -match '64') {

    $Response = Get-ChildItem -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\' | Get-ItemProperty | Select-Object -Property DisplayName,DisplayVersion | Where-Object {$_.Displayname -like "*$Software*"}
    $Result = @{
        Computer     = $IPAddress
        Processor    = $Env:PROCESSOR_ARCHITECTURE
        OSName       = [System.Environment]::OSVersion.VersionString
        OSArck       = (wmic os get osarchitecture)[2]
        Software     = $Response.DisplayName
        SoftArck     = '64-Bit'
        Version      = $Response.DisplayVersion
    }
}
else {

    $Result = @{
        Computer     = $IPAddress
        Processor    = $Env:PROCESSOR_ARCHITECTURE
        OSName       = [System.Environment]::OSVersion.VersionString
        OSArck       = (wmic os get osarchitecture)[2]
        Software     = 'N/A'
        SoftArck     = 'N/A'
        Version      = 'N/A'
    }
}

[PSCustomObject]$Result