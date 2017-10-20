function Move-Message
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$TargetFolderPath
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		if ($TargetFolderPath -ne $null)
		{
			$Folder = Get-FolderFromPath -FolderPath $TargetFolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		}
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $ItemURI + "/move"
			$MoveItemPost = "{`"DestinationId`": `"" + $Folder.Id + "`"}"
			write-host $MoveItemPost
			return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $MoveItemPost
		}
	}
}
