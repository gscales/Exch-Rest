function Set-EXRUnPinEmail
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,

		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$ItemRESTURI
	)
	Process
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$Props = Get-EXRPinnedEmailProperty
		$Props[0].Value = "null"
		$Props[1].Value = "null"
		return Update-Message -MailboxName $MailboxName -ItemURI $ItemRESTURI -ExPropList $Props -AccessToken $AccessToken
		
	}
}
