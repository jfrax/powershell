
#has 14 members before fix
Set-DynamicDistributionGroup -Identity "All Alachua" -RecipientFilter {(City -eq 'Alachua') -and ((Company -eq 'RTI Biologics') -or (Company -eq 'RTI Surgical') -or (Company -eq 'RTI Donor Services'))}

#has 0 members before fix
Set-DynamicDistributionGroup -Identity "All Austin" -RecipientFilter {(City -eq 'Austin') -and ((Company -eq 'RTI Biologics') -or (Company -eq 'RTI Surgical') -or (Company -eq 'RTI Donor Services'))}

#has 14 members before fix
Set-DynamicDistributionGroup -Identity "All Greenville" -RecipientFilter {(City -eq 'Greenville') -and ((Company -eq 'RTI Biologics') -or (Company -eq 'RTI Surgical') -or (Company -eq 'RTI Donor Services'))}

#has 16 members before fix
Set-DynamicDistributionGroup -Identity "All Houten" -RecipientFilter {(City -eq 'Houten') -and ((Company -eq 'RTI Biologics') -or (Company -eq 'RTI Surgical') -or (Company -eq 'RTI Donor Services'))}

#has 3 members before fix
Set-DynamicDistributionGroup -Identity "All Marquette" -RecipientFilter {(City -eq 'Marquette') -and ((Company -eq 'RTI Biologics') -or (Company -eq 'RTI Surgical') -or (Company -eq 'RTI Donor Services'))}

#has 0 members before fix
Set-DynamicDistributionGroup -Identity "All Raleigh" -RecipientFilter {(City -eq 'Raleigh') -and ((Company -eq 'RTI Biologics') -or (Company -eq 'RTI Surgical') -or (Company -eq 'RTI Donor Services'))}

#has 3 members before fix
Set-DynamicDistributionGroup -Identity "All RTI Leadership" -RecipientFilter {(DirectReports -ne $null) -and (Co -eq 'United States') -and ( -not (HiddenFromAddressListsEnabled -eq 'True'))}

#has 141 members before fix
Set-DynamicDistributionGroup -Identity "All RTIX - Germany" -RecipientFilter {(CountryOrRegion -eq 'Germany') -and ((Company -eq 'Tutogen Medical GmbH') -or (Company -eq 'RTI Surgical GmbH') -or (Company -eq 'RTI Surgical'))}

#has 16 members before fix
Set-DynamicDistributionGroup -Identity "All RTIX - US" -RecipientFilter {(Co -eq 'United States') -and ((Company -eq 'RTI Biologics') -or (Company -eq 'RTI Surgical'))}

#has 164 members before fix
Set-DynamicDistributionGroup -Identity "All RTIX - International" -RecipientFilter {(-not(Co -eq 'United States')) -and ((Company -eq 'RTI Biologics') -or (Company -eq 'RTI Surgical') -or (Company -eq 'RTI Surgical GmbH') -or (Company -eq 'Tutogen Medical GmbH'))}
