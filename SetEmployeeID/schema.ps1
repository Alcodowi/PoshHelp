$obj = New-Object -Type PSCustomObject
$obj | Add-Member -Type NoteProperty -Name "Anchor-ObjectGUID|Guid" -Value "00000000-0000-0000-0000-000000000001"
$obj | Add-Member -Type NoteProperty -Name "objectClass|String" -Value "Externaluser"
$obj | Add-Member -Type NoteProperty -Name "DistinguishedName|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "SamAccountName|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "employeeType|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "msExchExtensionCustomAttribute1|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "mail|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "Domain|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "PDCEmulator|String" -Value ""
$obj 