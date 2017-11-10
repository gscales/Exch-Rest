function Get-ProfiledToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName
	)
	Process
	{
		$HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
		if ($Script:TokenCache.ContainsKey($HostDomain))
		{
			return $Script:TokenCache[$HostDomain]
		}
	}
}