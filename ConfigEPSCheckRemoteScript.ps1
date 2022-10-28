$IPAddress = $($args[0])
Set-Location -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\

$File = (Get-ChildItem -Path . | Select-Object -ExpandProperty Name)[-10..-1]

try {
    $ForwardServer = (Select-String -Path $File -Pattern '\[\d+.\d+.\d+.\d+:.*LBManager.getConnect')[-1]
}
catch {
    $ForwardServer = "Error"
}
try {
    $FuelServer    = (Select-String -Path $File -Pattern 'IP:.*EpsService.checkIPPort')[-1]
}
catch {
    $FuelServer    = "Error"
}
try {
    $SiteID        = (Select-String -Path $File -Pattern 'request:/d/dwr/loginService/epsInit.*\[DWRService.execute\]')[-1]
}
catch {
    $SiteID        = "Error"
}
try {
    $RPOSPort      = (Select-String -Path $File -Pattern '\[\d+.\d+.\d+.\d+:.*RmiConnectFactory.makeObject')[-2]
}
catch {
    $RPOSPort      = "Error"
}


[PSCustomObject]@{

    IPAddress         = $IPAddress
    ForwardServerIP   = $ForwardServer
    FuelServerIP      = $FuelServer
    SiteID            = $SiteID
    RPOSPort          = $RPOSPort
}