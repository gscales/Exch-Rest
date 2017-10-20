function Get-ArchiveFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings/ArchiveFolder"
		$JsonObject = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		$folderId = $JsonObject.value.ToString()
		$HttpClient = Get-HTTPClient($MailboxName)
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$folderId')"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}
