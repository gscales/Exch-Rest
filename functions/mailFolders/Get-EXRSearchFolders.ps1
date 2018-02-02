function Get-EXRSearchFolders{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=1, Mandatory=$false)] [String]$FolderName
    )
    Process{

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
        $SearchFolder = Get-EXRWellKnownFolder -MailboxName $MailboxName -AccessToken $AccessToken -FolderName SearchFolders
        if([String]::IsNullOrEmpty($FolderName)){
            Get-EXRChildFolders -Folder $SearchFolder -MailboxName $MailboxName -AccessToken  $AccessToken 
        }
        else{
            $HttpClient = Get-HTTPClient -MailboxName $MailboxName
            $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL = $SearchFolder.FolderRestURI
            $RequestURL = $RequestURL += "/childfolders/?`$filter=DisplayName eq '$FolderName'"
            $tfTargetFolder = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			if ($tfTargetFolder.Value.displayname -match $FolderName)
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
