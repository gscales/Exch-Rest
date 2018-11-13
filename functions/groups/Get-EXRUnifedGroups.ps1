function Get-EXRUnifedGroups
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$GroupName,

		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$mail

	)
	Process
	{
		
		if($AccessToken -eq $null)
		{
			$AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
			if($AccessToken -eq $null){
				$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
			}                 
		}
		if([String]::IsNullOrEmpty($MailboxName)){
			$MailboxName = $AccessToken.mailbox
		}  
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$RequestURL = Get-EndPoint -AccessToken $AccessToken -Segment "/groups?`$filter=groupTypes/any(c:c+eq+'Unified')"
		if (![String]::IsNullOrEmpty($GroupName))
		{
			$RequestURL = Get-EndPoint -AccessToken $AccessToken -Segment "/groups?`$filter=displayName eq '$GroupName'"
		}
		if($mail){
			$RequestURL = Get-EndPoint -AccessToken $AccessToken -Segment "/groups?`$filter=mail eq '$mail'"
		}
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value)
			{
				Write-Output $Message
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
	}
}
