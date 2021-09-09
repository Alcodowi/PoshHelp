param (
    $Username,
	$Password,
    $Credentials,
	$OperationType,
    [bool] $usepagedimport,
	$pagesize
    )

$DebugFilePath = "C:\scripts\EXOImport.txt"

if(!(Test-Path $DebugFilePath))
    {
        $DebugFile = New-Item -Path $DebugFilePath -ItemType File
    }
    else
    {
        $DebugFile = Get-Item -Path $DebugFilePath
    }
    
"Starting Import as: " + $OperationType + " " + (Get-Date) | Out-File $DebugFile -Append
$Credentials = New-Object System.Management.Automation.PSCredential $Username,$Password

Import-Module AzureADPreview
$aad = Connect-AzureAD -Credential $Credentials

if ($aad.Environment.Name -eq "AzureCloud"){
     "Authenticated to Azure AD " + (Get-Date) | out-file $DebugFile -Append
    "     Retreiving AAD Users "  + (Get-Date)  | out-file $DebugFile -Append 
    $AADUsers = Get-AzureADUser -filter "userType eq 'Member' and accountEnabled eq true" -All $true     

    if ($AADUsers){
        "     $($AADUsers.count) retreived from AAD "  + (Get-Date) | out-file $DebugFile -Append
    }
}

if (!$Global:mailboxes -and ($AADUsers)) {
    # Create Remote PowerShell Session to Exchange Online
    try {
        if ($global:EXOSession){Remove-PSSession $global:EXOSession.Id}
        "Opening a new EXO RPS Session "  + (Get-Date)   | out-file $DebugFile -Append
        $Global:EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credentials -Authentication Basic -AllowRedirection -ErrorVariable $EXOError 
        Import-PSSession $Global:EXOSession -AllowClobber
        "     Opened a new RPS Session "  + (Get-Date)  | out-file $DebugFile -Append
        "     EXO Session $($EXOSession.Id) $($EXOSession.ConfigurationName) $($EXOSession.State) "  + (Get-Date)  | Out-File $DebugFile -Append
    }
    catch {
     
         "    Failed creating a new RPS Session " + (Get-Date) | out-file $DebugFile -Append
         $EXOError | out-file $DebugFile -Append
         
    }

    "Retreiving Mailboxes from EXO " +(Get-Date) | out-file $DebugFile -Append
    $Global:mailboxes = Get-Mailbox -ResultSize Unlimited 
    "    Found $($mailboxes.count) mailboxes found "+ (Get-Date) | Out-File $DebugFile -Append
    [int]$Global:intMailboxes = $mailboxes.count

    "     Processing mailboxes " +(Get-Date) | Out-File $DebugFile -Append
    # Process Users without Mailboxes 
    Foreach ($User in $AADUsers) {
        $AADMailbox = $mailboxes | Where-Object {($_.UserPrincipalName -eq $User.UserPrincipalName)}         
        # No Mailbox, just AADUser
        if (!$AADMailbox){
                $mailboxObj = @{}
                $mailboxObj.add("objectGuid", $User.ObjectId)        
                $mailboxObj.add("objectClass", "MailUser")    
                $mailboxObj.add("IsMailboxEnabled", $false) # Boolean
                $mailboxObj.Add("AADUserPrincipalName",$User.userPrincipalName)
                $mailboxObj.Add("AADAccountEnabled",$User.accountEnabled)
                $mailboxObj.Add("AADDirSyncEnabled",$User.dirSyncEnabled)
                $mailboxObj.Add("AADDisplayName",$User.displayName)
                $mailboxObj.Add("AADGivenName",$User.givenName)
                $mailboxObj.Add("AADImmutableId",$User.immutableId)
                $mailboxObj.Add("AADLastDirSyncTime",[string]$User.lastDirSyncTime)
                $mailboxObj.Add("AADMail",$User.mail)
                $mailboxObj.Add("AADMailNickname",$User.mailNickname)
                try{
                    if ($User.onPremisesSecurityIdentifier) {
                           # Create SID .NET object using SID string from AAD S-1-500-........ 
                            $sid = New-Object system.Security.Principal.SecurityIdentifier $User.onPremisesSecurityIdentifier
                    
                            #Create a byte array for the length of the users SID
                            $BinarySid = new-object byte[]($sid.BinaryLength)
                            #Copy the binary sid into the byte array, starting at index 0
                            $sid.GetBinaryForm($BinarySid, 0)
                            $mailboxObj.Add("AADonPremiseSID",$BinarySid)    
                        }
                }
                Catch{
                   "ERROR: $_.Exception.Message" | Out-File $DebugFile -Append  
                }
                if ($User.proxyAddresses)
                {
                    $proxyAddresses = @()
                    foreach($address in $User.proxyAddresses) {
                       $proxyAddresses += $address
                    }
                    $mailboxObj.Add("AADProxyAddresses",($proxyAddresses))
                }

                $mailboxObj.Add("AADSurname",$User.surname)
                $mailboxObj.Add("AADTelephoneNumber",$User.telephoneNumber) 
                $mailboxObj.Add("AADPasswordPolicies",$User.passwordPolicies)           
                if ($AADUser.showInAddressList){$mailboxObj.Add("AADShowInAddressList",$User.showInAddressList)}
                $mailboxObj.Add("AADCompanyName",$User.companyName)
                $mailboxObj.Add("AADCountry",$User.country)
                $mailboxObj.Add("AADPhysicalDeliveryOfficeName",$User.physicalDeliveryOfficeName)   
                $mailboxObj.Add("AADUsageLocation",$User.usageLocation)
                $mailboxObj.Add("AADJobTitle",$User.jobTitle)
                $mailboxObj.Add("AADMobile",$User.mobile)  
                $mailboxObj.Add("AADSipProxyAddress",$User.sipProxyAddress)

                if ($User.otherMails)
                  {  
                    $otherMails = @()
                    foreach($otheraddress in $User.otherMails) {
                       $otherMails += $otheraddress
                    }
                    $mailboxObj.Add("AADOtherMails",($otherMails))
                  }                                  
                $mailboxObj.Add("AADCity",$User.city)
               $mailboxObj
        }
    }

    # Process Users with Mailboxes
    if ($Global:mailboxes) {
        foreach ($mailbox in $mailboxes){
            # Get AADUser Object from AADUsers Collection retreived earlier
            $AADUser = $AADUsers | Where-Object {($_.UserPrincipalName -eq $mailbox.UserPrincipalName)}            
            
            if ($AADUser){ 
                # Create Object to pass to MA
                # Mailbox Attrs
                $mailboxObj = @{}
                $mailboxObj.add("objectGuid", $AADUser.ObjectId)        
                $mailboxObj.add("objectClass", "MailUser")        
                $mailboxObj.add("accountName", $mailbox.SamAccountName)        
                $mailboxObj.add("LitHold", $mailbox.LitigationHoldEnabled)
                $mailboxObj.add("IsDirSynced", $mailbox.IsDirSynced)
                $mailboxObj.add("Database", $mailbox.Database)
                $mailboxObj.add("MailboxRegion", $mailbox.MailboxRegion)
                $mailboxObj.add("UseDatabaseRetentionDefaults", $mailbox.UseDatabaseRetentionDefaults)                        
                $mailboxObj.add("RetainDeletedItemsUntilBackup", $mailbox.RetainDeletedItemsUntilBackup)
                $mailboxObj.add("RetentionHoldEnabled", $mailbox.RetentionHoldEnabled)
                $mailboxObj.add("EndDateForRetentionHold", [string]$mailbox.EndDateForRetentionHold)
                $mailboxObj.add("StartDateForRetentionHold", [string]$mailbox.StartDateForRetentionHold)
                $mailboxObj.add("LitigationHoldDate", [string]$mailbox.LitigationHoldDate)
                $mailboxObj.add("ComplianceTagHoldApplied", $mailbox.ComplianceTagHoldApplied) #[boolean]
                $mailboxObj.add("DelayHoldApplied", $mailbox.DelayHoldApplied)
                $mailboxObj.add("InactiveMailboxRetireTime", [string]$mailbox.InactiveMailboxRetireTime)
                $mailboxObj.add("OrphanSoftDeleteTrackingTime", [string]$mailbox.OrphanSoftDeleteTrackingTime)
                $mailboxObj.add("LitigationHoldDuration", $mailbox.LitigationHoldDuration)
                $mailboxObj.add("NeedOnlineMailbox", $True)

                # Set the number of days for Litigation Hold Duration if not unlimited 
                if ($mailbox.LitigationHoldDuration){
                    $days = $mailbox.LitigationHoldDuration
                    If (!$days.Equals("Unlimited")){
                        $days = $days.Split(".")                        
                        $mailboxObj.add("LitigationHoldDays", $days[0])
                    }
                }

                $mailboxObj.add("RetentionPolicy", $mailbox.RetentionPolicy)
                $mailboxObj.add("ExchangeGuid", [string]$mailbox.ExchangeGuid)             

                if ($mailbox.MailboxLocations)
                    {
                        $MailboxLocations = @()
                        foreach($location in $mailbox.MailboxLocations) {
                           $MailboxLocations += $location
                        }
                        $mailboxObj.Add("MailboxLocations",($MailboxLocations))
                    }

                $mailboxObj.add("ExchangeSecurityDescriptor", $mailbox.ExchangeSecurityDescriptor) 
                $mailboxObj.add("ExchangeUserAccountControl", $mailbox.ExchangeUserAccountControl) 
                $mailboxObj.add("ForwardingAddress", $mailbox.ForwardingAddress)         
                $mailboxObj.add("ForwardingSmtpAddress", $mailbox.ForwardingSmtpAddress)
                $mailboxObj.add("RetainDeletedItemsFor", $mailbox.RetainDeletedItemsFor) 
                $mailboxObj.add("IsMailboxEnabled", $mailbox.IsMailboxEnabled) # Boolean

                if ($mailbox.Languages)
                {
                    $Languages = @()
                    foreach($lang in $mailbox.Languages) {
                       $Languages += $lang.Name
                    }
                    $mailboxObj.Add("Languages",($Languages))
                }

                $mailboxObj.add("IsLinked", $mailbox.IsLinked)  # Boolean
                $mailboxObj.add("IsShared", $mailbox.IsShared) # Boolean
                $mailboxObj.add("ResourceType", $mailbox.ResourceType) 
                $mailboxObj.add("SamAccountName", $mailbox.SamAccountName) 
                $mailboxObj.add("ServerLegacyDN", $mailbox.ServerLegacyDN) 
                $mailboxObj.add("ServerName", $mailbox.ServerName) 
                $mailboxObj.add("UserPrincipalName", $mailbox.UserPrincipalName) 
                $mailboxObj.add("UMEnabled", $mailbox.UMEnabled) # Boolean
                $mailboxObj.add("WindowsLiveID", $mailbox.WindowsLiveID) 
                $mailboxObj.add("MicrosoftOnlineServicesID", $mailbox.MicrosoftOnlineServicesID) 
                $mailboxObj.add("RoleAssignmentPolicy", $mailbox.RoleAssignmentPolicy) 
                $mailboxObj.add("MailboxPlan", $mailbox.MailboxPlan) 
                $mailboxObj.add("ArchiveDatabase", $mailbox.ArchiveDatabase) 
                $mailboxObj.add("ArchiveGuid", [string]$mailbox.ArchiveGuid) 

                if ($mailbox.ArchiveName)
                {
                    $Archives = @()
                    foreach($archive in $mailbox.ArchiveName) {
                       $Archives += $archive
                    }
                    $mailboxObj.Add("ArchiveName",($Archives))
                }

                $mailboxObj.add("ArchiveStatus", $mailbox.ArchiveStatus) 
                $mailboxObj.add("ArchiveState", $mailbox.ArchiveState) 
                $mailboxObj.add("MailboxMoveTargetMDB", $mailbox.MailboxMoveTargetMDB) 
                $mailboxObj.add("MailboxMoveSourceMDB", $mailbox.MailboxMoveSourceMDB) 
                $mailboxObj.add("MailboxMoveFlags", $mailbox.MailboxMoveFlags) 
                $mailboxObj.add("MailboxMoveRemoteHostName", $mailbox.MailboxMoveRemoteHostName) 
                $mailboxObj.add("MailboxMoveBatchName", $mailbox.MailboxMoveBatchName) 
                $mailboxObj.add("MailboxMoveStatus", $mailbox.MailboxMoveStatus) 
                $mailboxObj.add("MailboxRelease", $mailbox.MailboxRelease) 
                $mailboxObj.add("WhenMailboxCreated", [string]$mailbox.WhenMailboxCreated) 
                $mailboxObj.add("UsageLocation", $mailbox.UsageLocation) 
                $mailboxObj.add("IsSoftDeletedByRemove", $mailbox.IsSoftDeletedByRemove)  #[boolean]
                $mailboxObj.add("IsSoftDeletedByDisable", $mailbox.IsSoftDeletedByDisable)  #[boolean]
                $mailboxObj.add("IsInactiveMailbox", $mailbox.IsInactiveMailbox)  #[boolean]
                $mailboxObj.add("WhenSoftDeleted", [string]$mailbox.WhenSoftDeleted) 
        
                if ($mailbox.InPlaceHolds)
                {
                    $Holds = @()
                    foreach($hold in $mailbox.InPlaceHolds) {
                       $Holds += $hold
                    }
                    $mailboxObj.Add("InPlaceHolds",($Holds))
                }

                $mailboxObj.add("AccountDisabled", $mailbox.AccountDisabled) #[boolean]
                $mailboxObj.add("HasPicture", $mailbox.HasPicture) #[boolean]
                $mailboxObj.add("Alias", $mailbox.Alias) 
                $mailboxObj.add("OrganizationalUnit", $mailbox.OrganizationalUnit) 

                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute1", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute1", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute2", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute2", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute3", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute3", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute4", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute4", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute5", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute5", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute6", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute6", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute7", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute7", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute8", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute8", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute9", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute9", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute10", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute10", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute11", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute11", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute12", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute12", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute13", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute13", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute14", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute14", $null)}
                if ($mailbox.CustomAttribute1){$mailboxObj.add("CustomAttribute15", $mailbox.CustomAttribute1)}else{$mailboxObj.add("CustomAttribute15", $null)}

               if ($mailbox.ExtensionCustomAttribute1)
                {
                    $ExtensionCustomAttribute1 = @()
                    foreach($val in $mailbox.ExtensionCustomAttribute1) {
                       $ExtensionCustomAttribute1 += $val
                    }
                    $mailboxObj.Add("ExtensionCustomAttribute1",($ExtensionCustomAttribute1))
                }

                 if ($mailbox.ExtensionCustomAttribute2)
                {
                    $ExtensionCustomAttribute2 = @()
                    foreach($val in $mailbox.ExtensionCustomAttribute2) {
                       $ExtensionCustomAttribute2 += $val
                    }
                    $mailboxObj.Add("ExtensionCustomAttribute2",($ExtensionCustomAttribute2))
                }

                 if ($mailbox.ExtensionCustomAttribute3)
                {
                    $ExtensionCustomAttribute3 = @()
                    foreach($val in $mailbox.ExtensionCustomAttribute3) {
                       $ExtensionCustomAttribute3 += $val
                    }
                    $mailboxObj.Add("ExtensionCustomAttribute3",($ExtensionCustomAttribute3))
                }

                 if ($mailbox.ExtensionCustomAttribute4)
                {
                    $ExtensionCustomAttribute4 = @()
                    foreach($val in $mailbox.ExtensionCustomAttribute4) {
                       $ExtensionCustomAttribute4 += $val
                    }
                    $mailboxObj.Add("ExtensionCustomAttribute4",($ExtensionCustomAttribute4))
                }

                 if ($mailbox.ExtensionCustomAttribute5)
                {
                    $ExtensionCustomAttribute5 = @()
                    foreach($val in $mailbox.ExtensionCustomAttribute5) {
                       $ExtensionCustomAttribute5 += $val
                    }
                    $mailboxObj.Add("ExtensionCustomAttribute5",($ExtensionCustomAttribute5))
                }

                $mailboxObj.add("DisplayName", $mailbox.DisplayName) 

                if ($mailbox.EmailAddresses)
                {
                    $Addresses = @()
                    foreach($Address in $mailbox.EmailAddresses) {
                       $Addresses += $Address
                    }
                    $mailboxObj.Add("EmailAddresses",($Addresses))
                }

                $mailboxObj.add("ExternalDirectoryObjectId", $mailbox.ExternalDirectoryObjectId) 
                $mailboxObj.add("HiddenFromAddressListsEnabled", $mailbox.HiddenFromAddressListsEnabled)  #[boolean]
                $mailboxObj.add("LegacyExchangeDN", $mailbox.LegacyExchangeDN) 
                $mailboxObj.add("LastExchangeChangedTime", [string]$mailbox.LastExchangeChangedTime) 
                $mailboxObj.add("ModerationEnabled", $mailbox.ModerationEnabled)   #[boolean]
            
                if ($mailbox.PoliciesIncluded)
                {
                    $PoliciesIncluded = @()
                    foreach($InPolicy in $mailbox.PoliciesIncluded) {
                       $PoliciesIncluded += $InPolicy
                    }
                    $mailboxObj.Add("PoliciesIncluded",($PoliciesIncluded))
                }

                if ($mailbox.PoliciesExcluded)
                {
                    $PoliciesExcluded = @()
                    foreach($ExPolicy in $mailbox.PoliciesExcluded) {
                       $PoliciesExcluded += $ExPolicy
                    }
                    $mailboxObj.Add("PoliciesExcluded",($PoliciesExcluded))
                }

                $mailboxObj.add("EmailAddressPolicyEnabled", $mailbox.EmailAddressPolicyEnabled)   #[boolean]
                $mailboxObj.add("PrimarySmtpAddress", $mailbox.PrimarySmtpAddress) 
                $mailboxObj.add("RecipientType", $mailbox.RecipientType) 
                $mailboxObj.add("RecipientTypeDetails", $mailbox.RecipientTypeDetails) 
                $mailboxObj.add("WindowsEmailAddress", $mailbox.WindowsEmailAddress) 
                $mailboxObj.add("Identity", $mailbox.Identity) 
                $mailboxObj.add("ExchangeVersion", $mailbox.ExchangeVersion) 
                $mailboxObj.add("Name", $mailbox.Name) 
                $mailboxObj.add("DistinguishedName", $mailbox.DistinguishedName) 
                $mailboxObj.add("Guid", [string]$mailbox.Guid) 
                $mailboxObj.add("WhenChanged", [string]$mailbox.WhenChanged) 
                $mailboxObj.add("WhenCreatedUTC", [string]$mailbox.WhenCreatedUTC) 
                $mailboxObj.add("ObjectState", $mailbox.ObjectState) 
                
                # AAD User Attrs
                $mailboxObj.Add("AADUserPrincipalName",$AADUser.userPrincipalName)
                $mailboxObj.Add("AADAccountEnabled",$AADUser.accountEnabled)
                $mailboxObj.Add("AADDirSyncEnabled",$AADUser.dirSyncEnabled)
                $mailboxObj.Add("AADDisplayName",$AADUser.displayName)
                $mailboxObj.Add("AADGivenName",$AADUser.givenName)
                $mailboxObj.Add("AADImmutableId",$AADUser.immutableId)
                $mailboxObj.Add("AADLastDirSyncTime",[string]$AADUser.lastDirSyncTime)
                $mailboxObj.Add("AADMail",$AADUser.mail)
                $mailboxObj.Add("AADMailNickname",$AADUser.mailNickname)
                try{
                    if ($AADUser.onPremisesSecurityIdentifier) {
                           # Create SID .NET object using SID string from AAD S-1-500-........ 
                            $sid = New-Object system.Security.Principal.SecurityIdentifier $AADUser.onPremisesSecurityIdentifier
                    
                            #Create a byte array for the length of the users SID
                            $BinarySid = new-object byte[]($sid.BinaryLength)

                            #Copy the binary sid into the byte array, starting at index 0
                            $sid.GetBinaryForm($BinarySid, 0)
                            $mailboxObj.Add("AADonPremiseSID",$BinarySid)    
                        }
                }
                Catch{
                   "ERROR: $_.Exception.Message" | Out-File $DebugFile -Append  
                }
                if ($AADUser.proxyAddresses)
                {
                    $proxyAddresses = @()
                    foreach($address in $AADUser.proxyAddresses) {
                       $proxyAddresses += $address
                    }
                    $mailboxObj.Add("AADProxyAddresses",($proxyAddresses))
                }

                $mailboxObj.Add("AADSurname",$AADUser.surname)
                $mailboxObj.Add("AADTelephoneNumber",$AADUser.telephoneNumber) 
                $mailboxObj.Add("AADPasswordPolicies",$AADUser.passwordPolicies)
                if ($AADUser.showInAddressList){$mailboxObj.Add("AADShowInAddressList",$AADUser.showInAddressList)}
                $mailboxObj.Add("AADCompanyName",$AADUser.companyName)
                $mailboxObj.Add("AADCountry",$AADUser.country)
                $mailboxObj.Add("AADPhysicalDeliveryOfficeName",$AADUser.physicalDeliveryOfficeName)   
                $mailboxObj.Add("AADUsageLocation",$AADUser.usageLocation)
                $mailboxObj.Add("AADJobTitle",$AADUser.jobTitle)
                $mailboxObj.Add("AADMobile",$AADUser.mobile)  
                $mailboxObj.Add("AADSipProxyAddress",$AADUser.sipProxyAddress)

                if ($AADUser.otherMails)
                  {  
                    $otherMails = @()
                    foreach($otheraddress in $AADUser.otherMails) {
                       $otherMails += $otheraddress
                    }
                    $mailboxObj.Add("AADOtherMails",($otherMails))
                  }           
                       
                $mailboxObj.Add("AADCity",$AADUser.city)
     
                # Pass the User Object to the MA
                $mailboxObj 
        }
      }
    }
}

"Completed Import " + (Get-Date) | Out-File $DebugFile -Append 