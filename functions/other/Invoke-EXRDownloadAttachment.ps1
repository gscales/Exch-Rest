function Invoke-EXRDownloadAttachment
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$AttachmentURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-EXRHTTPClient -MailboxName $MailboxName
		$AttachmentURI = $AttachmentURI + "?`$expand"
		$AttachmentObj = Invoke-EXRRestGet -RequestURL $AttachmentURI -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -TrackStatus:$true
		return $AttachmentObj
	}
}
