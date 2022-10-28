$Site = $($args[0])
$File = (Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\ | Select-Object -ExpandProperty FullName)[-10..-1]


try {
    $ForwardServer = [String](Select-String -Path $File -Pattern '\[\d+.\d+.\d+.\d+:.*LBManager.getConnect' | Select-String -Pattern '192' -NotMatch)[-1] | Select-String -Pattern '\d+.\d+.\d+.\d+:\d+' | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
}
catch {
    $ForwardServer = "ManualCheck"
}


try {
    $FuelServer    = [String](Select-String -Path $File -Pattern 'IP:.*EpsService.checkIPPort')[-1] | Select-String -Pattern '\d+.\d+.\d+.\d+ port:\d+' | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
}
catch {
    $FuelServer    = "ManualCheck"
}


try {
    $SiteID = [String](Select-String -Path $File -Pattern 'request:/d/dwr/loginService/epsInit.*\[DWRService.execute\]')[-1] | Select-String -Pattern "[^0-9]$Site[^0-9]" | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value

}
catch {
    $SiteID        = "ManualCheck"
}


# try {
#     $RPOSPort      = ([String](Select-String -Path $File -Pattern '\[\d+.\d+.\d+.\d+:.*RmiConnectFactory.makeObject')[-1] | Select-String -Pattern ':[0-9]{5}]'  | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value).Replace(":","").Replace("]","")
# }
# catch {
#     $RPOSPort      = "ManualCheck"
# }


[PSCustomObject]@{

    ForwardServerIP   = $ForwardServer.Split(':')[0]
    ForwardServerPort = $ForwardServer.Split(':')[1]
    FuelServerIP      = $FuelServer.Split(' ')[0]
    FuelServerPort    = $FuelServer.Split(' ')[1].Split(':')[1]
    SiteID            = $SiteID.Replace("'","")
#   RPOSPort          = $RPOSPort
}