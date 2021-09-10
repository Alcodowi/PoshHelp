param
(
	$username,
	$password,
 	$ExportType 
)

begin
{   
    # // Logging
    $DebugFilePath = "C:\scripts\EXOExport.txt"
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
           if ($can -eq "AA_needOnlineMailbox"){$AA_needOnlineMailbox = $_."AA_needOnlineMailbox"}                  
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
            $nickname = $mvUserObj.Attributes.mailNickname.Values.valuestring
            $mboxEnabled = $mvUserObj.Attributes.ExchangeMailboxEnabled.Values.valueBoolean
            $AA_needOnlineMailbox = $mvUserObj.Attributes.AA_needOnlineMailbox.Values.valueBoolean
            
            ## If Alias in changed Attrs and user not mailbox enabled then Enable 
            if (!$mboxEnabled -and $AA_needOnlineMailbox -eq $true){
                 "Enable user locally for O365 Mailbox"
                "Enable Remote Mailbox" | Out-File $DebugFile -Append
                $userAccount | Out-File $DebugFile -Append
                $nickname | Out-File $DebugFile -Append
                $localremoteMbx = Enable-RemoteMailbox -Identity $userAccount -RemoteRoutingAddress "$($nickname)@jbdlab.mail.onmicrosoft.com"
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
