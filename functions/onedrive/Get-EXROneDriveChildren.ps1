function Get-EXROneDriveChildren
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$DriveRESTURI,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$FolderPath
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-EXRHTTPClient -MailboxName $MailboxName
		if ([String]::IsNullOrEmpty($DriveRESTURI))
		{
			$EndPoint = Get-EXREndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $FolderPath + ":/children"
		}
		else
		{
			$RequestURL = $DriveRESTURI + "/children"
		}
		$JSONOutput = Invoke-EXRRestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		foreach ($Item in $JSONOutput.value)
		{
			Add-Member -InputObject $Item -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EXREndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $Item.Id + "')")
			write-output $Item
		}
		
	}
}
