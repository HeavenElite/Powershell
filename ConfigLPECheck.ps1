$Computers = Import-Csv -Path .\ITLab\ITLabData.csv | Where-Object {$_.LPE -eq 'LPEServer'}

for ($i=0; $i -lt ($Computers | Measure-Object).Count; $i++) {

    $Username   = $Computers[$i].Account
    $Password   = ConvertTo-SecureString -AsPlainText -Force $Computers[$i].Password
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password

    $LoopContinue = $true

    try {

        $Response = Invoke-Command -ComputerName $Computers[$i].IP -Credential $Credential -ScriptBlock {Get-Content -Path "C:\Retalix\LPE\CPTCPServer.exe.config"} -ErrorAction Stop
    }

    catch {

        $Report = @{
        
            Site          = $Computers[$i].Site
            Type          = $Computers[$i].Type
            Environment   = $Computers[$i].Test
            IPAddress     = $Computers[$i].IP
            Chain         = 'Offline'
            Branch        = 'Offline'
            WebServiceUrl = 'Offline'
        }
        
        $Report | Select-Object -Property Site,Type,Environment,IPAddress,Chain,Branch,WebServiceUrl | Format-Table -AutoSize
        $Report | Select-Object -Property Site,Type,Environment,IPAddress,Chain,Branch,WebServiceUrl | Export-Csv -Path ".\ConfigLPECheckReport-$(Get-Date -Format 'yyyy.MM.dd').csv" -Append -NoTypeInformation
    
        $LoopContinue = $false
    }

    finally {

    if ($LoopContinue) {

        $Result   = ([xml]$Response).configuration.SAFServer.add | Where-Object {$_.key -eq 'Chain' -or $_.key -eq 'Branch' -or $_.key -eq 'WebServiceUrl'}
        
        $Report = @{
            
            Site          = $Computers[$i].Site
            Type          = $Computers[$i].Type
            Environment   = $Computers[$i].Test
            IPAddress     = $Computers[$i].IP
            Chain         = $Result[0].value
            Branch        = $Result[1].value
            WebServiceUrl = $Result[2].value
        }
        
        $Report | Select-Object -Property Site,Type,Environment,IPAddress,Chain,Branch,WebServiceUrl | Format-Table -AutoSize
        $Report | Select-Object -Property Site,Type,Environment,IPAddress,Chain,Branch,WebServiceUrl | Export-Csv -Path ".\ConfigLPECheckReport-$(Get-Date -Format 'yyyy.MM.dd').csv" -Append -NoTypeInformation
    }
    }
}

Import-Csv -Path ".\ConfigLPECheckReport-$(Get-Date -Format 'yyyy.MM.dd').csv" | Sort-Object -Property Environment,Site | Format-Table -AutoSize

[System.Console]::Beep(1000,1000)