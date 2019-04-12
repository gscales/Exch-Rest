function Invoke-EXRDownloadAttachment
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$AttachmentURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken,

		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$DownloadPath
	)
	Begin
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
		$AttachmentURI = $AttachmentURI + "?`$expand"
		$AttachmentObj = Invoke-RestGet -RequestURL $AttachmentURI -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -TrackStatus:$true
		if([String]::IsNullOrEmpty($DownloadPath)){
			return $AttachmentObj
		}else{
			 $attachBytes = [System.Convert]::FromBase64String($AttachmentObj.ContentBytes)
			 [System.IO.File]::WriteAllBytes($DownloadPath,$attachBytes) 
		}

	}
}
