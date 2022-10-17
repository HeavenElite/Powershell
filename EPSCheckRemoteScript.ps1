if (Test-Path -Path C:\Users\$env:UserName\AppData\Roaming\Shell\EPS\uninstall.exe -PathType Leaf) {

    Set-Location C:\Users\$env:UserName\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\

    if ((Get-ChildItem . | Select-Object -ExpandProperty Name | Measure-Object -line | Select-Object -ExpandProperty Lines) -ne 1 -or 0) {

        $FileList = Get-ChildItem . | Select-Object -ExpandProperty Name
        ((Select-String -Path $FileList[-1] 'eps version' | Select-Object -ExpandProperty Line)[-1].Substring(36,10)) -replace 'V|"',''

    }

    elseif ((Get-ChildItem . | Select-Object -ExpandProperty Name | Measure-Object -line | Select-Object -ExpandProperty Lines) -eq 1) {

        $File = Get-ChildItem . | Select-Object -ExpandProperty Name
        ((Select-String -Path $File 'eps version' | Select-Object -ExpandProperty Line)[-1].Substring(36,10)) -replace 'V|"',''

    }

    elseif ((Get-ChildItem . | Select-Object -ExpandProperty Name | Measure-Object -line | Select-Object -ExpandProperty Lines) -eq 0) {

        Write-Output "Log file has not been generated, check it manually by RDP. `n"

    }

    else {

        Write-Output "There must be something strange happeded. `n"

    }
}

else {
    
    Write-Output 'NotInstalled'

}