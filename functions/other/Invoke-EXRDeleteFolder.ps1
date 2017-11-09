function Invoke-EXRDeleteFolder
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
		$FolderPath
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
			$confirmation = Read-Host "Are you Sure You Want To proceed with deleting Folder"
			if ($confirmation -eq 'y')
			{
				$HttpClient = Get-HTTPClient -MailboxName $MailboxName
				$RequestURL = $Folder.FolderRestURI
				return Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			}
			else
			{
				Write-Host "skipped deletion"
			}
		}
		
		
	}
}
