$Index = $($args[0])
$Site  = $($args[1])
$IPAddress = $($args[2])


$FilePath  = (Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\ | Select-Object -ExpandProperty FullName)[$Index]
$FileName  = (Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\ | Select-Object -ExpandProperty Name)[$Index]

Copy-Item -Path $FilePath -Destination 'C:\'
$FilePath  = "C:\$FileName"


$FTPServer = "ftp://Username:Password@192.168.0.141/$Site-$IPAddress-$FileName"
$URL = New-Object System.Uri($FTPServer)
$WebClient = New-Object System.Net.WebClient
$WebClient.UploadFile($URL, $FilePath)

Remove-Item -Path $FilePath
Write-Output "$Site-$IPAddress-$FileName"