if (Test-Path -Path C:\Users\$env:UserName\AppData\Roaming\Shell\EPS\uninstall.exe -PathType Leaf) {

    Set-Location C:\Users\$env:UserName\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\
    $FileList = Get-ChildItem . | Select-Object -ExpandProperty Name
    (Select-String -Path $FileList[-1] 'eps version' | Select-Object -ExpandProperty Line)[-1].Substring(36,10)
}
else {
    Write-Output 'EPS NotInstalled'
}