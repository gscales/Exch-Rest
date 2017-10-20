function Invoke-DownloadAttachment
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
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$AttachmentURI = $AttachmentURI + "?`$expand"
		$AttachmentObj = Invoke-RestGet -RequestURL $AttachmentURI -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		return $AttachmentObj
	}
}
