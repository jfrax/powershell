Import-Module Sqlps -DisableNameChecking;

# Invoke-Sqlcmd -Query "select guid from Logins where Active = 1 and guid is not null" -ServerInstance "rtix-sqap-01" -Database "chart"