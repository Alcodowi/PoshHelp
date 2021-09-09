param (
    $Username,
    $Password,
    $Credentials,
    $OperationType,
    $NBDomain = "ADDomain",
    $MailboxName = "newstaffreports@customer.com.au",
    $pageupsender = "reports-out@saas.provider.com",
    $downloadDirectory = "\\fileserver\appshare$\SaaSReports",
    $EWSURI=[system.URI] "https://webmail.customer.com.au/ews/Exchange.asmx",
    $ProcessedInboxFolder = "Processed",
    [bool] $usepagedimport,
    $pagesize
    )


[string]$CSVFile = $null 

## Load EWS Managed API dll
# https://www.microsoft.com/en-us/download/details.aspx?id=42951
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

## Set Exchange Version (Exchange2010, Exchange2010_SP1, Exchange2010_SP2, Exchange2013 or Exchange2013_SP1)
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1

## Create Exchange Service Object
$ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

## Set Credentials (using Username and Password passed from the MA Config
$creds = New-Object System.Net.NetworkCredential($Username,$Password,$NBDomain) 
$ExchangeService.Credentials = $creds  
## Set AutoDiscover URL from Params
$ExchangeService.Url = $EWSURI  

#Bind to the Inbox folder
$SearchFilterAttachments = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$InboxFolderID= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,$InboxFolderID)  

## Connect to the mailbox, find the attachements 
$numItems = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
$InboxResults = $Inbox.FindItems($SearchFilterAttachments,$numItems)

[int]$attachmentint = 0
$pageupexport = $false

# Processes by the most recent received message first
foreach($MailItem in $InboxResults.Items){
	$MailItem.Load()
	foreach($attach in $MailItem.Attachments){

        # the first attachment should be the one we're looking for (the most recent received)
        # but we'll check to make sure it is from our sender and the attachment name matches what we're expecting
		if ($attachmentint = 0 -or !$pageupexport)
            {
                if ($MailItem.Sender.Address -contains $pageupsender)
                    {                    
                   	   # make sure the attachment is the CSV file we're after and not a signature graphic
                       if ($attach.Name.ToString() -like 'IT_Setup*.csv')
                        {
                            # Extract the CSV file and save to the File Share
                            $attach.Load()
                            $File = new-object System.IO.FileStream(($downloadDirectory + “\” + $attach.Name.ToString()), [System.IO.FileMode]::Create)
		                    $File.Write($attach.Content, 0, $attach.Content.Length)
		                    $File.Close()
                            
                            $CSVFile = $downloadDirectory + “\” + $attach.Name.ToString()
                            $pageupexport = $true
                        }
                    }
            }
               
       $attachmentint = $attachmentint + 1
	}
}

#Move the emails with attachments
if ($attachmentint -gt 0)
    {      
        #Get the ID of the folder to move to  
        $Folder =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)  
        $Folder.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
        $FolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$ProcessedInboxFolder)
        $ProcessedFolder = $Inbox.FindFolders($FolderSearchFilter,$Folder)  
  
        #Define ItemView to retreive just 100 Items    
        $numItems =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)  
        $AttachedFiles = $null    
        do
            {    
                $AttachedFiles = $Inbox.FindItems($SearchFilterAttachments,$numItems)   
                foreach($Attachment in $AttachedFiles.Items)
                    {      
                        # Move the Message  
                        $Attachment.Move($ProcessedFolder.Folders[0].Id)  
                    }    
                $numItems.Offset += $AttachedFiles.Items.Count    
            }
        while($AttachedFiles.MoreAvailable -eq $true) 
    }


## Make sure we have a file to import
if ($CSVFile)
    {
        #We have a new file
        $importfile = $true 
    }
    Else
    {
        ## We don't have a new file from the Inbox. We have to locate the last file we put in the fileshare
        $newestfiledate = [DateTime]::MinValue
        $newestfiledate | Out-File $DebugFile -Append 
        # Find files in the fileshare

        get-childitem $downloadDirectory | Where {$_.name -like 'IT_Setup*.csv'} | ForEach-Object {
                if ($_.LastWriteTime -gt $newestfiledate)# -and -not $_.PSIsContainer) 
                {
                       # Check the file properties                     
                       $newestfiledate = $_.LastWriteTime                     
                       $LastFileinDir = $_.Name
                       $LastFileinDirDate = $_.LastWriteTime
                       $importfile = $true

                       $CSVFile =  $downloadDirectory + “\” + $LastFileinDir                  
                }
        }
   }


# Parse the CSV first, then do logic on it in the next step
If ($importfile -and $CSVFile){
     $importdata = Import-Csv $CSVFile | foreach {      
        New-Object PSObject -prop @{
            StaffID = $_.'Sign On';
            Surname = $_.'Applicant last name';
            GivenName = $_.'Applicant first name';
            PreferredName = $_.'Preferred name';
            MiddleName = $_.'Middle name';
            State = $_.'Home state/territory';
            Title = $_.'Position Title';
            EmployeeType = $_.'Employment type';
            Company = $_.Business;
            Department = $_.Department;
            BirthDate = $_.'Date of birth'
            StartDate = $_.'Date Started'
        }
    }
}

# Perform validation on the data and bring through the new staff accordingly
 ForEach($user in $importdata)
        {
         #Only bring in users we have a StaffID (Anchor for) 
         if ($user.StaffID)
            {
                # Only bring in users where their StartDate is within the next 7 days
                $todaysDate = get-date              
                if ($user.StartDate) { $userStartDate = get-date($user.StartDate)
                    }
                    if ($userStartDate.AddDays(-7) -lt $todaysDate.AddDays(7)) 
                        {
                        $obj = @{}
                        $obj.Add("StaffID",$user.StaffID)
                        $obj.Add("objectClass", "user")
                        $obj.Add("Surname",$user.Surname)
                        $obj.Add("GivenName",$user.GivenName)
                   
                        $obj.Add("MiddleName",$user.MiddleName)
                        $obj.Add("State",$user.State)
                        $obj.Add("Title",$user.Title)
                        $obj.Add("Company",$user.Company)
                        $obj.Add("Department",$user.Department)
                
                        if ($user.BirthDate) {
                             $birthDate =  Get-Date ($user.BirthDate)
                             [string]$birthDateFormat = $birthDate.ToString("yyyy-MM-dd")
                             $obj.Add("BirthDate",$birthDateFormat)
                        }
                                        
                        if ($user.StartDate) {
                             $startDate =  Get-Date ($user.StartDate)
                             [string]$startDateFormat = $startDate.ToString("yyyy-MM-dd")
                             $obj.Add("StartDate",$startDateFormat)
                        }

                        if ($user.PreferredName) {
                            $obj.Add("PreferredName",$user.PreferredName)
                        }
                        else 
                        {
                            if ($user.GivenName) {
                                 $obj.Add("PreferredName",$user.GivenName)
                            }
                        }

                        # Pass the object to the MA
                        $obj
                    }
            }
}
#endregion