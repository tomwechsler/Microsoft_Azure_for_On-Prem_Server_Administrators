#Enable AD Recycle Bin
Enable-ADOptionalFeature `
-Identity "Recycle Bin Feature" `
-Scope ForestOrConfigurationSet `
-Target "master.pri" `
-Confirm:$False


#Did it work?
Get-ADOptionalFeature -Identity "Recycle Bin Feature"

#The sample
Get-ADUser -Identity frabets

Get-ADUser -Identity frabets | Remove-ADUser -Confirm:$False

Get-ADUser -Identity frabets 

Get-ADObject -Filter {Name -like "frabets*"} -IncludeDeletedObjects

Get-ADObject -Filter {Name -like "frabets*"} `
-IncludeDeletedObjects | Restore-ADObject

Get-ADObject -Filter {Name -like "frabets"}
