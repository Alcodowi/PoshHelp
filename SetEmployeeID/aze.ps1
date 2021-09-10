$domains ="Alco.local","jbdlab.local"
$alco="Alco.local"
$jbdlab="jbdlab.local"
foreach ($Domain in $Domains){

    $GetDom = Get-ADDomain $Domain
    $PDC = $GetDom.pdcemulator
    #$users += Get-ADUser -filter {EmployeeID -NOTLIKE '*'} -Server $PDC
    #$users += Get-ADUser -filter {EmployeeID -NOTLIKE '*' -AND displayname -like '*ext*' -AND enabled -eq 'true' -and employeeType -eq "External"} -Server $PDC -Properties givenName,sn,mail,displayname,name,DistinguishedName,employeeType,msExchExtensionCustomAttribute1 | 

    $users += Get-ADUser -filter {EmployeeID -NOTLIKE '*' -AND enabled -eq 'true'} -Server $PDC -Properties givenName,sn,mail,displayname,name,DistinguishedName,employeeType | 
    where-object {
    ($_.samaccountname -notlike 'prd.*') -and
    ($_.samaccountname -notlike 'svc.*') -and
    ($_.samaccountname -notlike 'srv.*') -and
    ($_.samaccountname -notlike 'sec.*') -and
    ($_.samaccountname -notlike '*.adm') -and
    ($_.samaccountname -notlike '*-adm') -and
    ($_.samaccountname -notlike 'SM*') -and
    ($_.samaccountname -notlike '*HealthMailbox*') -and
    #($_.sn -like '*-ext') -and
    ($_.sn -like '*') -and
    ($_.UserPrincipalName -notlike '*.local') -and
    ($_.description -notlike '*service account*') -and
    ($_.givenName -notlike '*service*') -and
    ($_.sn -notlike '*service*') -and
    ($_.displayname -notlike '*service*')  -and
    ($_.DistinguishedName -notlike 'CN=Users,DC=Alco,DC=local')
    ($_.DistinguishedName -notlike 'CN=monitoring*')
    } 
}