function Get-ADdirectReports
{
    PARAM ($SamAccountName)
    Get-Aduser -identity $SamAccountName -Properties directreports | %{
        $_.directreports | ForEach-Object -Process {
            # Output the current Object information
            Get-ADUser -identity $Psitem -Properties mail,manager | Select-Object -Property Name, SamAccountName, Mail, @{ L = "Manager"; E = { (Get-Aduser -iden $psitem.manager).samaccountname } }
            
            # Find the DirectReports of the current item ($PSItem / $_)
            Get-ADdirectReports -SamAccountName $PSItem
        }
    }
}