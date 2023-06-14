﻿connect-ExchangeOnline

Test-ApplicationAccessPolicy -Identity $emailToBeBlocked -AppId $AppID
Test-ApplicationAccessPolicy -Identity $emailToBeAllowed -AppId $AppID


New-ApplicationAccessPolicy -AppId $AppID -PolicyScopeGroupId $PolicyScopeGroup -AccessRight RestrictAccess -Description "Restrict this app to members of distribution group EvenUsers."

Test-ApplicationAccessPolicy -Identity $emailToBeBlocked -AppId $AppID
Test-ApplicationAccessPolicy -Identity $emailToBeAllowed -AppId $AppID


# Disconnect from Exchange
Disconnect-ExchangeOnline