$obj = New-Object -Type PSCustomObject
$obj | Add-Member -Type NoteProperty -Name "Anchor-SamAccountName|String" -Value "00000000"
$obj | Add-Member -Type NoteProperty -Name "objectClass|String" -Value "Externaluser"
$obj | Add-Member -Type NoteProperty -Name "DistinguishedName|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "employeeType|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "msExchExtensionCustomAttribute1|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "mail|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "Domain|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "PDCEmulator|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "name|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "employeeID|String" -Value ""
$obj | Add-Member -Type NoteProperty -Name "NeedID|Boolean" -Value ""
$obj | Add-Member -Type NoteProperty -Name "FromPowershell|Boolean" -Value ""
$obj 