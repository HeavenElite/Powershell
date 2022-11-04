$Computers = Import-Csv .\ITLab\ITLabData.csv
# | Where-Object {$_.IP -like '192.168.1.132' -or $_.IP -like '192.168.0.32' -or $_.IP -like '192.168.1.62' -or $_.IP -like '192.168.0.92' -or $_.IP -like '192.168.0.52'}
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
            ID = 'Offline'
            Status = 'Offline' 
            Suggestion = 'CheckHW'
        }

        [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','ID','Status','Suggestion' | Format-Table -AutoSize
        [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','ID','Status','Suggestion' | Export-Csv -Path $Path -Append -NoTypeInformation
    
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
                ID = 'Idle'
                Status = 'Idle' 
                Suggestion = 'PleaseLogin'
            }
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','ID','Status','Suggestion' | Format-Table -AutoSize
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','ID','Status','Suggestion' | Export-Csv -Path $Path -Append -NoTypeInformation
        }

        elseif (($Response | Measure-Object).Count -gt 1) {

            for ($m=1; $m -lt $Response.Length; $m++) {
                $Report = @{

                    Site = $Computers[$i].Site
                    IPAddress = $Computers[$i].IP
                    Account = $Response[$m].Split(' ')[2]
                    Session = $Response[$m].Split(' ')[1] -replace 'rdp-tcp#','RDP:' -replace 'console','Console'
                    ID = $Response[$m].Split(' ')[3]
                    Status = $Response[$m].Split(' ')[4]
                    Suggestion = 'ContactUser'
                }
            
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','ID','Status','Suggestion' | Format-Table -AutoSize
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','ID','Status','Suggestion' | Export-Csv -Path $Path -Append -NoTypeInformation
            }
        }

        else {

            $Report = @{

                Site = $Computers[$i].Site
                IPAddress = $Computers[$i].IP
                Account = 'Abnormal'
                Session = 'Abnormal'
                ID = 'Abnormal'
                Status = 'Abnormal'
                Suggestion = 'Abnormal'
            }
        
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','ID','Status','Suggestion' | Format-Table -AutoSize
            [PSCustomObject]$Report | Select-Object -Property 'Site','IPAddress','Account','Session','ID','Status','Suggestion' | Export-Csv -Path $Path -Append -NoTypeInformation
        }
    }
    }
}

Import-Csv -Path ".\$Path" | Sort-Object -Property Suggestion,Site | Format-Table -AutoSize

[System.Console]::Beep(1000,1000)