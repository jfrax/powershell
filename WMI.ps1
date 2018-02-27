$query = "select * from Win32_Bios"
$query = "select * from Win32_Product"
$query = "select * from Win32_OperatingSystem"
$query = "select * from Win32_SoftwareFeature"


Get-WmiObject -Query $query | select-object lastuse | Format-table -autosize