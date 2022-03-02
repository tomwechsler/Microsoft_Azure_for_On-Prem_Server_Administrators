$da = (Get-ADDomain).domainsid
$da = $da.tostring() + "-500"
Get-ADUser -Identity $da
