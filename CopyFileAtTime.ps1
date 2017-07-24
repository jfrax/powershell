$targetDate = Get-Date "7/5/2017 02:00 AM"
$sourceFile = "C:\staging\Site.Master"
$destFile = "\\apps.rtix.com\web\Production\HelpDesk\Site.Master"

while ((Get-Date) -lt $targetDate) {
    $timeLeft = (NEW-TIMESPAN –Start (Get-Date) –End $targetDate);
    Write-Host ('{0} Days, {1} Hours, {2} Minutes remaining' -f $timeLeft.Days, $timeLeft.Hours, $timeLeft.Minutes)
    Start-Sleep -s 60

}

Copy-Item $sourceFile $destFile
Write-Host 'Done!';