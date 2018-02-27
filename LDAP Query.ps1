get-aduser -SearchBase "OU=IT,OU=Users,OU=Tutogen,DC=rtix,DC=com" -Filter * -Properties SamAccountName, c -ResultSetSize 5000 | Select SamAccountName, c

get-aduser -SearchBase "OU=Tutogen Medical GmbH,OU=Users,OU=Tutogen,DC=rtix,DC=com" -Filter * -Properties SamAccountName, c -ResultSetSize 5000 | Select SamAccountName, c

get-aduser -SearchBase "OU=NEU-Users2,OU=Users,OU=Tutogen,DC=rtix,DC=com" -Filter * -Properties SamAccountName, c -ResultSetSize 5000 | Select SamAccountName, c

get-aduser -SearchBase "OU=RTI Surgical GmbH,OU=Users,OU=Tutogen,DC=rtix,DC=com" -Filter * -Properties SamAccountName, c -ResultSetSize 5000 | Select SamAccountName, c