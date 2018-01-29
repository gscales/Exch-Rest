function Get-EXRFolderFromId
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,

		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$FolderId
	)
	Begin
	{
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
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders('" + $FolderId + "')"
		$PropList  = @()
		$PropList += Get-EXRFolderPath 
		 if($PropList -ne $null){
            $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
            $RequestURL += "`?`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            
        }
		$tfTargetFolder = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		if($tfTargetFolder -ne $null){
			$folderId = $tfTargetFolder.Id.ToString()
			Add-Member -InputObject $tfTargetFolder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
			Expand-ExtendedProperties -Item $tfTargetFolder
			return, $tfTargetFolder	
		}else{return,$null}
	}
}
