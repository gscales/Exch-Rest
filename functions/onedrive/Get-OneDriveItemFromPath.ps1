function Get-OneDriveItemFromPath
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$OneDriveFilePath
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $OneDriveFilePath
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Add-Member -InputObject $JSONOutput -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $JSONOutput.Id + "')")
		return $JSONOutput
	}
}
