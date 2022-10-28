$IPAddress = $($args[0])
$Site      = $($args[1])
$File      = (Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\ | Select-Object -ExpandProperty FullName)[-10..-1]


try {

    $ForwardServer = ((Select-String -Path $File -Pattern '\[\d+.\d+.\d+.\d+:.*LBManager.getConnect' | Select-String -Pattern '192' -NotMatch)[-1] | Select-String -Pattern '\d+.\d+.\d+.\d+:\d+').Matches.Value

}
catch {
    $ForwardServer = "Error"
}


try {
    $FuelServer    = ((Select-String -Path $File -Pattern 'IP:.*EpsService.checkIPPort')[-1] | Select-String -Pattern '\d+.\d+.\d+.\d+ port:\d+').Matches.Value
}
catch {
    $FuelServer    = "Error"
}


try {

    $SiteID = ((Select-String -Path $File -Pattern 'request:/d/dwr/loginService/epsInit.*\[DWRService.execute\]' | Select-String -Pattern "[^0-9]$Site[^0-9]")[-1] | Select-String "'[0-9]{4}'").Matches.Value

}
catch {
    $SiteID        = "Error"
}


try {
    $RPOSPort      = ((Select-String -Path $File -Pattern '\[\d+.\d+.\d+.\d+:.*RmiConnectFactory.makeObject')[-1] | Select-String -Pattern ':[0-9]{5}]').Matches.Value
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