param (
    $Username,
	$Password,
    $Credentials,
	$OperationType,
    [bool] $usepagedimport,
	$pagesize
    )

$DebugFilePath = "C:\scripts\Employeeid\ADImport.txt"

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
$domains ="Alco.local"
foreach ($Domain in $Domains){

    $GetDom = Get-ADDomain $Domain
    $PDC = $GetDom.pdcemulator

    <#$user += Get-ADUser -filter {EmployeeID -NOTLIKE '*' -AND displayname -like '*ext*' -AND enabled -eq 'true' -and employeeType -eq "External"} -Server $Domain -Properties givenName,sn,mail,displayname,name,DistinguishedName,employeeType,msExchExtensionCustomAttribute1 | 
    where-object {
    ($_.samaccountname -notlike 'prd.*') -and
    ($_.samaccountname -notlike 'svc.*') -and
    ($_.samaccountname -notlike 'srv.*') -and
    ($_.samaccountname -notlike 'sec.*') -and
    ($_.samaccountname -notlike '*.adm') -and
    ($_.samaccountname -notlike '*-adm') -and
    #($_.sn -like '*-ext') -and
    ($_.sn -like '*') -and
    ($_.UserPrincipalName -notlike '*.group') -and
    ($_.description -notlike '*service account*') -and
    ($_.givenName -notlike '*service*') -and
    ($_.sn -notlike '*service*') -and
    ($_.displayname -notlike '*service*')  
    }
    #>
    $users += Get-ADUser -filter {EmployeeID -NOTLIKE '*'} -Server $PDC
}
    "     $($Users.count) users retreived from Active Directory "  + (Get-Date) | out-file $DebugFile -Append
    

"     Processing user " +(Get-Date) | Out-File $DebugFile -Append
# Process Users without Mailboxes 
    Foreach ($user in $users) { 
        $UserObj = @{}
        $UserObj.add("SID", $User.SID)
        $UserObj.add("objectClass", "Externaluser") 
        $UserObj.add("SamAccountName", $User.SamAccountName)
        $UserObj.add("PDCEmulator", $PDC)
        $UserObj.add("DistinguishedName",$User.DistinguishedName)
        $UserObj.add("employeeType",$User.employeeType)
        $UserObj.add("mail",$User.mail)
        $UserObj  
    }   
"Completed Import " + (Get-Date) | Out-File $DebugFile -Append 