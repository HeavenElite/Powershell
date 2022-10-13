function Remove-Backup {
    [CmdletBinding()]
    $Result = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {Remove-Item 'D:\*' -Recurse -Force -Confirm:$false}
    $Result
}
function Clear-Backup {
    [CmdletBinding()]
    $Result = Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {wbadmin.exe delete catalog -quiet}
    $Result
}
function Set-Backup {
    Invoke-Command -ComputerName $IPAddress -Credential $Credential -ScriptBlock {wbadmin.exe start backup -backupTarget:D: -include:C: -Quiet}
}
