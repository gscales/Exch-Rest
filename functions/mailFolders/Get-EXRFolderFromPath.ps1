function Get-EXRFolderFromPath
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken,

		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
		$PropList
	)
	process
	{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		if($AccessToken -eq $null)
		{
			$AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
			if($AccessToken -eq $null){
				$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
			}                 
		}
		if([String]::IsNullOrEmpty($MailboxName)){
			$MailboxName = $AccessToken.mailbox
		}  
		if($FolderPath.ToLower() -eq "mailboxroot"){
		    Get-EXRRootMailFolder -MailboxName $MailboxName -AccessToken $AccessToken
		}
		else{
			$HttpClient = Get-HTTPClient -MailboxName $MailboxName
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/msgfolderroot/childfolders?"
			#  $RootFolder = Invoke-EXRRestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			#Split the Search path into an array  
			$tfTargetFolder = $RootFolder
			$fldArray = $FolderPath.Split("\")
			#Loop through the Split Array and do a Search for each level of folder 
			for ($lint = 1; $lint -lt $fldArray.Length; $lint++)
			{
				#Perform search based on the displayname of each folder level
				$FolderName = $fldArray[$lint];
				$RequestURL = $RequestURL += "`$filter=DisplayName eq '$FolderName'"
        		if($PropList -ne $null){
           			 $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
           			 $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
       			}
				$tfTargetFolder = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
				if ($tfTargetFolder.Value.displayname -match $FolderName)
				{
					$folderId = $tfTargetFolder.value.Id.ToString()
					$RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$folderId')/childfolders?"
				}
				else
				{
					throw ("Folder Not found")
				}
			}
			if ($tfTargetFolder.Value -ne $null)
			{
				$folderId = $tfTargetFolder.Value.Id.ToString()
				Add-Member -InputObject $tfTargetFolder.Value -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
				Expand-ExtendedProperties -Item $tfTargetFolder.Value
				return, $tfTargetFolder.Value
			}
			else
			{
				throw ("Folder Not found")
			}
		}
	}
}
