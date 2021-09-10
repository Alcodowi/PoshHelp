param
(
	$username,
	$password,
 	$ExportType 
)

begin
{   
    # // Logging
    $DebugFilePath = "C:\scripts\Employeeid\EXOExport.txt"
    if(!(Test-Path $DebugFilePath))
        {$DebugFile = New-Item -Path $DebugFilePath -ItemType File}
    else
        {$DebugFile = Get-Item -Path $DebugFilePath}
    "Starting Export : " + (Get-Date) | Out-File $DebugFile -Append 
    "ExportType : $ExportType " | Out-File $DebugFile -Append

    # Password from the MA
    $securestring = ConvertTo-SecureString -AsPlainText $Password -Force
       
    Import-Module lithnetMIISAutomation
    $EXOMA = 'EmployeeID' 

}

process
{
    $error.clear() 
    $errorstatus = "success"
    $errordetails = $null 

    $Identifier = $_.'[Identifier]'
    $objectGuid = $_.'[DN]'
	
    "=========="  | Out-File $DebugFile -Append   
    "Changed Attributes" | Out-File $DebugFile -Append   
    $_.'[ChangedAttributeNames]' | Out-File $DebugFile -Append   

    #Loop through changes and update parameters
    foreach ($can in $_.'[ChangedAttributeNames]')
      {    
           if ($can -eq "employeeID"){$employeeID = $_."employeeID"}                  
      }
    
    #Supported ChangeType is Replace
    if ($_.'[ObjectModificationType]' -eq 'Add'){
        # Not doing add's just updates
    }

    #Supported ChangeType is Replace
    if ($_.'[ObjectModificationType]' -eq 'Replace')
  	{
            "Object Modification Type - Replace" | Out-File $DebugFile -Append   
            $_ | Out-File $DebugFile -Append   
                       
            $errorstatus = "success" 
            $userAccount = $_.AADUserPrincipalName 
            $csUserObj = Get-CSObject -DN $objectGuid -MA $EXOMA
            $userMVGuid = $csUserObj.MvGuid.Guid
            $mvUserObj = Get-MVObject -ID $csUserObj.MvGuid.Guid
            $displayName = $mvUserObj.Attributes.displayName.Values.valuestring
            $alias = $mvUserObj.Attributes.mailNickname.Values.valuestring
            $identity = $mvUserObj.Attributes.accountName.Values.valuestring
            $server = $mvUserObj.attributes.AA_pdcemulator
            $employeeID = $mvUserObj.attributes.employeeID
            
            ## If Alias in changed Attrs and user not mailbox enabled then Enable 
            if ($employeeID){
                 "Enable user locally for O365 Mailbox"
                "Enable Remote Mailbox" | Out-File $DebugFile -Append
                Set-ADUser -Identity $identity -Server $server -EmployeeID $employeeID
                # Enable Archive
                #"Enable Remote Mailbox In-Place Archive" | Out-File $DebugFile -Append   
                #$localremoteArchive = Set-RemoteMailbox -Identity $userAccount -ArchiveName "In-Place Archive - $($displayName)"                
            }  

        }					
    #Return the result to the MA    
    $obj = @{}
    $obj.Add("[Identifier]",$Identifier) 
    $obj.Add("[ErrorName]","success")  
    if($errordetails){$obj.Add("[ErrorDetail]",$errordetails) }  
    $obj
 }

end
{
   #All done                  
   "Completed Export : " + (Get-Date) | Out-File $DebugFile -Append 
} 
