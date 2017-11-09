function Invoke-EXRDeleteItem
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
		$ItemURI
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$confirmation = Read-Host "Are you Sure You Want To proceed with deleting the Item"
		if ($confirmation -eq 'y')
		{
			$HttpClient = Get-HTTPClient -MailboxName $MailboxName
			$RequestURL = $ItemURI
			return Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		}
		else
		{
			Write-Host "skipped deletion"
		}
		
		
	}
}
