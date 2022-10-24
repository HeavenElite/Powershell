$Computers = Import-Csv .\ITLab\ITLabData.csv
$Date = Get-Date -Format 'yyyy.MM.dd'
$Path = "UserCheckReport-$Date.csv"

if (Test-Path -Path ".\$Path") {

    Remove-Item -Path .\$Path
    Write-Output "Previous file $Path is removed. `n"
}

else {

    Write-Output "Csv file $Path is generated. `n"
}

for ($i=0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText $Computers[$i].Password -Force
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    $LoopContinue = $true

    try {
        $Response = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {qwinsta.exe /server:localhost | Select-String -Pattern '会话','Session','rdp-tcp#','运行中','Active' | ForEach-Object {$_ -replace ' +',' ' -replace '运行中','Active'}} -ErrorAction Stop
    }
    catch {
        $Report = @{

            Site = $Computers[$i].Site
            IPAddress = $Computers[$i].IP
            Account = 'Offline'
            Session = 'Offline'
            Status = 'Offline' 
            Suggestion = 'CheckHW'
        }

        [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','Status','Suggestion' | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','Status','Suggestion' | Export-Csv -Path $Path -Append -NoTypeInformation
    
        $LoopContinue = $false
    }
    finally {

    if ($LoopContinue) {

        if (($Response | Measure-Object).Count -eq 1) {
            
            $Report = @{

                Site = $Computers[$i].Site
                IPAddress = $Computers[$i].IP
                Account = 'Idle'
                Session = 'Idle'
                Status = 'Idle' 
                Suggestion = 'PleaseLogin'
            }
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','Status','Suggestion' | Format-Table -AutoSize
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','Status','Suggestion' | Export-Csv -Path $Path -Append -NoTypeInformation
        }

        elseif (($Response | Measure-Object).Count -gt 1) {

            for ($m=1; $m -lt $Response.Length; $m++) {
                $Report = @{

                    Site = $Computers[$i].Site
                    IPAddress = $Computers[$i].IP
                    Account = $Response[$m].Split(' ')[2]
                    Session = $Response[$m].Split(' ')[1] -replace 'rdp-tcp#','RDP: ' -replace 'console','Console'
                    Status = $Response[$m].Split(' ')[4]
                    Suggestion = 'ContactUser'
                }
            
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','Status','Suggestion' | Format-Table -AutoSize
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','Status','Suggestion' | Export-Csv -Path $Path -Append -NoTypeInformation
            }
        }

        else {

            $Report = @{

                Site = $Computers[$i].Site
                IPAddress = $Computers[$i].IP
                Account = 'Abnormal'
                Session = 'Abnormal'
                Status = 'Abnormal'
                Suggestion = 'Abnormal'
            }
        
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','Status','Suggestion' | Format-Table -AutoSize
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','Status','Suggestion' | Export-Csv -Path $Path -Append -NoTypeInformation
        }
    }
    }
}

[System.Console]::Beep(1000,1000)