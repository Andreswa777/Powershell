connect-ExchangeOnline
$emailToBeBlocked = "%Email Address - user you dont want app to see%"$emailToBeAllowed = "%Email Address - user you want app to see%"$AppID = "%App ID of App reg to test%"$PolicyScopeGroup = "%Mail-Enabled Security group of all emails you want the app to see%"#Test to see if App reg can see any email address
Test-ApplicationAccessPolicy -Identity $emailToBeBlocked -AppId $AppID
Test-ApplicationAccessPolicy -Identity $emailToBeAllowed -AppId $AppID


New-ApplicationAccessPolicy -AppId $AppID -PolicyScopeGroupId $PolicyScopeGroup -AccessRight RestrictAccess -Description "Restrict this app to members of distribution group EvenUsers."

Test-ApplicationAccessPolicy -Identity $emailToBeBlocked -AppId $AppID
Test-ApplicationAccessPolicy -Identity $emailToBeAllowed -AppId $AppID


# Disconnect from Exchange
Disconnect-ExchangeOnline