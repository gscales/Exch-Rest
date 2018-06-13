function Get-ProfiledToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$ResourceURL
	)
	Process
	{
		if([String]::IsNullOrEmpty($ResourceURL)){
			$ResourceURL = "graph.microsoft.com" 
		}
		if($Script:TokenCache.ContainsKey($ResourceURL)){
			$AccessToken = $null
			if([String]::IsNullOrEmpty($MailboxName)){
				$firstToken = $Script:TokenCache[$ResourceURL].GetEnumerator() | select -first 1
				$AccessToken =  $firstToken.Value
			}
			else
			{
				$HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
				if ($Script:TokenCache[$ResourceURL].ContainsKey($HostDomain))
				{				
					$AccessToken = $Script:TokenCache[$ResourceURL][$HostDomain]
				}
			}
			if($AccessToken -ne $null){
				$MailboxName = $AccessToken.mailbox
				#Check for expired Token
				$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
				$expiry = $minTime.AddSeconds($AccessToken.expires_on)
				if ($expiry -le [DateTime]::Now.ToUniversalTime().AddMinutes(10))
				{
					if ([bool]($AccessToken.PSobject.Properties.name -match "refresh_token"))
					{
							write-host "Refresh Token"
							$AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
					}
					else
					{
						throw "App Token has expired"
					}
					
				}
				return $AccessToken
			}
			
		}

	}
}