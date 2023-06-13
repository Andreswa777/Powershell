connect-ExchangeOnline

Test-ApplicationAccessPolicy -Identity adm-andres@parkadvertising.co.za -AppId ac86f593-0a3d-4c09-b360-4ff9d1514a67
Test-ApplicationAccessPolicy -Identity New#@mediashop.co.za -AppId ac86f593-0a3d-4c09-b360-4ff9d1514a67


New-ApplicationAccessPolicy -AppId ac86f593-0a3d-4c09-b360-4ff9d1514a67 -PolicyScopeGroupId RoomResources@NahanaCG.onmicrosoft.com -AccessRight RestrictAccess -Description "Restrict this app to members of distribution group EvenUsers."

Test-ApplicationAccessPolicy -Identity adm-andres@parkadvertising.co.za -AppId ac86f593-0a3d-4c09-b360-4ff9d1514a67
Test-ApplicationAccessPolicy -Identity New#@mediashop.co.za -AppId ac86f593-0a3d-4c09-b360-4ff9d1514a67


# Disconnect from Exchange
Disconnect-ExchangeOnline