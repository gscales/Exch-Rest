function Update-EXRFolder
{
	[CmdletBinding()]
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
		$FolderPost
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$Folder = Get-EXRFolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient -MailboxName $MailboxName
			$RequestURL = $Folder.FolderRestURI
			$FolderPostValue = $FolderPost
			return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $FolderPostValue
			
		}
		
		
	}
}
