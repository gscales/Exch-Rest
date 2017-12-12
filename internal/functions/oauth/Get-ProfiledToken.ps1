function Get-ProfiledToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName
	)
	Process
	{
		if([String]::IsNullOrEmpty($MailboxName)){
			$firstToken = $Script:TokenCache.GetEnumerator() | select -first 1
			return $firstToken.Value
		}
		else
		{
			$HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
			if ($Script:TokenCache.ContainsKey($HostDomain))
			{				
				return $Script:TokenCache[$HostDomain]
			}
		}

	}
}