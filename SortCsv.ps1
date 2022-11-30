param(
    [Parameter(Mandatory)]
    [String]$Path
)

$Data = Import-Csv -Path $Path | Sort-Object -Property Computer

for ($i=0; $i -lt $Data.Length; $i++) {

    $Report = @{
        Computer       = $Data.Computer[$i]
        BackupItem     = $Data.BackupItem[$i]
        OldestVersion  = $Data.OldestVersion[$i]
        LatestVersion  = $Data.LatestVersion[$i]
        VersionNumber  = $Data.VersionNumber[$i]
        BackupLocation = $Data.BackupLocation[$i]
    }

    [PSCustomObject]$Report | Select-Object -Property Computer,BackupItem,OldestVersion,LatestVersion,VersionNumber,BackupLocation
    [PSCustomObject]$Report | Select-Object -Property Computer,BackupItem,OldestVersion,LatestVersion,VersionNumber,BackupLocation | Export-Csv -Path "$Path.Sort" -Append -NoTypeInformation

}