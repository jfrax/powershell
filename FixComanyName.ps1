
get-aduser -filter {Company -eq 'RTI Surgical, Inc' } | Set-ADUser -Company 'RTI Surgical';

# get-aduser -filter { Company -like '*Donor*' } -Properties 'Company'

