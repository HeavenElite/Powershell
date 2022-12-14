$Computers = Import-Csv .\ITLab\ITLabData.csv
$Date = Get-Date -Format 'yyyy.MM.dd'
$Path = "D:\Desktop\Powershell\RDPCheckReport-$Date.csv"

for ($i=0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    $LoopContinue = $true

    try {

        $Data = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-WinEvent -LogName 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational' | Where-Object {$_.ID -eq '1149' -and $_.TimeCreated -like "*$(Get-Date -Format 'MM/dd/yyyy')*"}} -ErrorAction Stop
    }
    catch {

        $Report = @{

            User      = 'PCOffline'
            Device    = $Computers[$i].IP
            Account   = 'PCOffline'
            Domain    = 'PCOffline'
            LoginTime = 'PCOffline'  
        }

        [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Export-Csv -Path $Path -Append -NoTypeInformation

        $LoopContinue = $false 
    }    
    finally {
        
    if ($LoopContinue) {

        for ($m=0; $m -lt $Data.Length; $m++) {
        
            $Report = @{
        
                User      = ($Data[$m].Message | Select-String -Pattern '[0-9]+.*').Matches.Value
                Device    = $Computers[$i].IP
                Account   = ($Data[$m].Message | Select-String -Pattern '用户: .*|User: .*').Matches.Value -replace '用户: ','' -replace 'User: ','' -replace "`r",''
                Domain    = ($Data[$m].Message | Select-String -Pattern '域: .*|Domain: .*').Matches.Value -replace '域: ','' -replace 'Domain: ','' -replace "`r",''
                LoginTime = $Data[$m].TimeCreated
                
            }
            [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Format-Table -AutoSize
            [PSCustomObject]$Report | Select-Object -Property User,Device,Account,Domain,LoginTime | Export-Csv -Path $Path -Append -NoTypeInformation

        }
    }
    }
}

[System.Console]::Beep(1000,1000)