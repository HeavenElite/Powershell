$File          = (Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Roaming\Shell\EPS\webapp\workspace\log\system\).FullName
$ForwardServer = (Select-String -Path $File -Pattern '\[\d+.\d+.\d+.\d+:.*LBManager.getConnect')[-1]
$FuelServer    = (Select-String -Path $File -Pattern 'IP:.*EpsService.checkIPPort')[-1]
$SiteID        = (Select-String -Path $File -Pattern 'StoreID="\d+"').Matches.Value[-1]
$RPOSPort      = (Select-String -Path $File -Pattern '\[\d+.\d+.\d+.\d+:.*RmiConnectFactory.makeObject')[-2]

[PSCustomObject]@{

    ForwardServerIP   = $ForwardServer.Split(':')[0]
    ForwardServerPort = $ForwardServer.Split(':')[1]
    FuelServerIP      = $FuelServer[0].Split(':')[1]
    FuelServerPort    = $FuelServer[1].Split(':')[1]
    SiteID            = $SiteID.Split('"')[1]
    RPOSPort          = $RPOSPort
}