param (
    $Username,
	$Password,
    $Credentials,
	$OperationType,
    [bool] $usepagedimport,
	$pagesize
    )

$DebugFilePath = "C:\scripts\Employeeid\"

if(!(Test-Path $DebugFilePath))
    {
        $DebugFile = New-Item -Path $DebugFilePath -ItemType File
    }
    else
    {
        $DebugFile = Get-Item -Path $DebugFilePath
    }
    
"Starting Import as: " + $OperationType + " " + (Get-Date) | Out-File $DebugFile -Append
#
    # Password from the MA
    $securestring = ConvertTo-SecureString -AsPlainText $Password -Force
    # Username from the MA
    $username
    # PS Credential with Username and password from MA
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username, $securestring
    #Domains to search
    #$Domains = Get-ADDomain

    "Exporting User... " + (Get-Date) | out-file $DebugFile -Append
    "     Retreiving User  "  + (Get-Date)  | out-file $DebugFile -Append 
>
#$users | Add-Member -Type NoteProperty -Name "PDCEmulator" -Force -Value ""    
$domains ="Alco.local","jbdlab.local"
$alco="Alco.local"
$jbdlab="jbdlab.local"
$lastid="EXT002000"
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
    ($_.samaccountname -notlike 'HealthMailbox*') -and
    #($_.sn -like '*-ext') -and
    ($_.sn -like '*') -and
    ($_.UserPrincipalName -notlike '*.local') -and
    ($_.description -notlike '*service account*') -and
    ($_.givenName -notlike '*service*') -and
    ($_.sn -notlike '*service*') -and
    ($_.displayname -notlike '*service*')  -and
    ($_.DistinguishedName -notlike '*CN=Users,DC=Alco,DC=local')
    } 
}
    "     $($users.count) users retreived from Active Directory "  + (Get-Date) | out-file $DebugFile -Append
    

"     Processing user " +(Get-Date) | Out-File $DebugFile -Append
# Process Users without Mailboxes 
    Foreach ($user in $users) { 
        <#
        $UserObj = @{}
        $UserObj.add("SamAccountName", $User.SamAccountName)
        $UserObj.add("objectClass", "Externaluser") 
        
        if ($user.DistinguishedName -contains "DC=Alco,DC=local"){

        $UserObj.add("PDCEmulator", $alco)
        }
        else {
            $UserObj.add("PDCEmulator", $jbdlab) 
        }
        $UserObj.add("DistinguishedName",$User.DistinguishedName)
        $UserObj.add("employeeType",$User.employeeType)
        $UserObj.add("mail",$User.mail)
        $UserObj.add("name",$User.mail)
        $UserObj 
         #>
        
         if ($user.DistinguishedName -like "*DC=Alco,DC=local*"){
            $server=$alco
            }
            else {
            $server=$jbdlab
            }
        $employeeID=$lastid+1
         Set-ADUser -Identity $user.DistinguishedName -EmployeeID $employeeID -Server $server
        $lastID=$employeeID
        Write-Host $employeeID
    }   
"Completed Import " + (Get-Date) | Out-File $DebugFile -Append 