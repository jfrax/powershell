$executives = @("rtabor", "cbuzzerd");
$execAttributeName = "extensionAttribute10";
$countryAttributeName = "extensionAttribute11";


Import-Module -Name ActiveDirectory -ErrorAction 'Stop';

#=======================================================
# Functions
#=======================================================
function Get-ADDirectReports
{
    PARAM ($SamAccountName)
    Get-Aduser -identity $SamAccountName -Properties directreports | %{
        $_.directreports | ForEach-Object -Process {
            
            Get-ADUser -identity $Psitem -Properties mail,manager | Select-Object -Property Name, SamAccountName, Mail, @{ L = "Manager"; E = { (Get-Aduser -iden $psitem.manager).samaccountname } }
            Get-ADDirectReports -SamAccountName $PSItem
        }
    }
}



#=======================================================
#Update customAttribute11 to be the Executive above each user
#=======================================================
foreach($e in $executives) {

    foreach($adu in (Get-ADDirectReports $e)) {

        # Write-Host $adu.SamAccountName
        Set-ADUser $adu.SamAccountName -Replace @{$execAttributeName = $e}
    }

}




#=======================================================
#Update customAttribute10 to be the Country of each user
#=======================================================

foreach ($adu in Get-ADUser -Filter 'c -like "*" -and manager -eq "CN=Ryan Tabor,OU=IT,OU=RTI Users,DC=rtix,DC=com"' -Properties "Country") {
    Set-ADUser $adu -Replace @{$countryAttributeName = $adu.Country}
}



