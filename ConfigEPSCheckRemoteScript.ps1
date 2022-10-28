$File          = (Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\).FullName
$ForwardServer = ([String](Select-String -Path $File[-1] -Pattern 'LBManager.getConnect')[0]).Split(' ')[1].Split('[')[1].Replace(']','')
$FuelServer    = ([String](Select-String -Path $File[-1] -Pattern 'IP:.*EpsService.checkIPPort')[0]).Split(' ')[2,3]
$SiteID        = (Select-String -Path $File -Pattern 'StoreID="\d+"').Matches.Value[-1]
$RPOSPort      = ([String](Select-String -Path $File[-1] -Pattern 'RmiConnectFactory.makeObject')[0]).Split('[')[2].Split(':')[1].Replace(']','')

[PSCustomObject]@{

    ForwardServerIP   = $ForwardServer.Split(':')[0]
    ForwardServerPort = $ForwardServer.Split(':')[1]
    FuelServerIP      = $FuelServer[0].Split(':')[1]
    FuelServerPort    = $FuelServer[1].Split(':')[1]
    SiteID            = $SiteID.Split('"')[1]
    RPOSPort          = $RPOSPort
}