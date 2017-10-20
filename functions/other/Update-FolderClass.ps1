function Update-FolderClass
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$FolderClass
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $Folder.FolderRestURI
			$UpdateFolderPost = "{`"SingleValueExtendedProperties`": [{`"PropertyId`":`"String 0x3613`",`"Value`":`"" + $FolderClass + "`"}]}"
			return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $UpdateFolderPost
			
		}
		
		
	}
}
