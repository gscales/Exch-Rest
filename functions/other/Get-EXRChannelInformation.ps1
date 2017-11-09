function Get-EXRChannelInformation
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$Group,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
		$Channel
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "groups"
		$RequestURL = $EndPoint + "('" + $Group.Id + "')/channels('" + $Channel.Id + "')"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		
		
	}
}
