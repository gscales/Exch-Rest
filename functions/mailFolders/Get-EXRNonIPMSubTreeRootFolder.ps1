function Get-EXRNonIPMSubTreeRootFolder
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
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
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Root"
		$PropList = @()
		$NProp = Get-EXRNamedProperty -DataType Binary -Guid "E49D64DA-9F3B-41AC-9684-C6E01F30CDFA" -Type String -Id TeamChatFolderEntryId
		$PropList += $NProp
		if($PropList -ne $null){
            $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
            $RequestURL += "`?`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            Write-Host $RequestURL
        }
		$tfTargetFolder = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		$folderId = $tfTargetFolder.Id.ToString()
		Add-Member -InputObject $tfTargetFolder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
		Expand-ExtendedProperties -Item $tfTargetFolder
		return, $tfTargetFolder
	}
}
