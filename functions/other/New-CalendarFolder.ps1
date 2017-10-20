function New-CalendarFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$DisplayName
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/calendars"
		$NewFolderPost = "{`"Name`": `"" + $DisplayName + "`"}"
		write-host $NewFolderPost
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
		
		
	}
}
