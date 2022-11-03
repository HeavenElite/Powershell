$Index = $($args[0])
$FilePath  = (Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\ | Select-Object -ExpandProperty FullName)[$Index]
$FileName  = (Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\ | Select-Object -ExpandProperty Name)[$Index]

$FTPServer = "ftp://Laurence:ShellLPE!23@192.168.0.141/$FileName"
$URL = New-Object System.Uri($FTPServer)
$WebClient = New-Object System.Net.WebClient
$WebClient.UploadFile($URL, $FilePath)

Write-Output "$FileName"