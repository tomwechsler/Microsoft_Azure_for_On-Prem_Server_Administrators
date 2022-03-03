Get-ADGroupMember "Domain Admins" | Format-Table Name

Get-ADGroupMember -Identity "Enterprise Admins" -Recursive

Get-ADGroupMember -Identity Administrators

Get-ADPrincipalGroupMembership Administrator | Format-Table Name

Get-ADUser -Filter 'memberof -recursivematch "CN=Domain Admins,CN=users,DC=prime,DC=pri"' | Format-Table Name
