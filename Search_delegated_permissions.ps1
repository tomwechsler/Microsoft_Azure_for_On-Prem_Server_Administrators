#More filters can be found here: http://www.ldapexplorer.com
$filter = "(|(objectClass=domain)(objectClass=organizationalUnit)(objectClass=group)(sAMAccountType=805306368)(objectCategory=Computer))" 

#("LDAP://DOMAINCONTROLLER/LDAP") Replace DOMAINCONTROLLER AND LDAP with your values
$bSearch = New-Object System.DirectoryServices.DirectoryEntry("LDAP://DC01/DC=zodiac,DC=local") 
$dSearch = New-Object System.DirectoryServices.DirectorySearcher($bSearch)
$dSearch.SearchRoot = $bSearch
$dSearch.PageSize = 1000
$dSearch.Filter = $filter
$dSearch.SearchScope = "Subtree"

#List of extended permissions: https://technet.microsoft.com/en-us/library/ff405676.aspx
$extPerms = `
        '00299570-246d-11d0-a768-00aa006e0529', #reset password
        '0'

$results = @()

foreach ($objResult in $dSearch.FindAll())
{
    $obj = $objResult.GetDirectoryEntry()

    Write-Host "Searching... " $obj.distinguishedName

    $permissions = $obj.PsBase.ObjectSecurity.GetAccessRules($true,$false,[Security.Principal.NTAccount])
    
    $results += $permissions | Where-Object { `
            $_.AccessControlType -eq 'Allow' -and ($_.ObjectType -in $extPerms) -and $_.IdentityReference -notin ('NT AUTHORITY\SELF', 'NT AUTHORITY\SYSTEM', 'S-1-5-32-548') `
            } | Select-Object `
        @{n='Object'; e={$obj.distinguishedName}}, 
        @{n='Account'; e={$_.IdentityReference}},
        @{n='Permission'; e={$_.ActiveDirectoryRights}}

}

#The output directly in Out-GridView
$results | Out-GridView
