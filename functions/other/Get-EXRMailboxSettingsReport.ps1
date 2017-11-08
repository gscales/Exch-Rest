function Get-EXRMailboxSettingsReport
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[psobject]
		$Mailboxes,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$CertFileName,
		
		[Parameter(Mandatory = $True)]
		[Security.SecureString]
		$password
	)
	Begin
	{
		$rptCollection = @()
		$AccessToken = Get-EXRAppOnlyToken -CertFileName $CertFileName -password $password
		$HttpClient = Get-EXRHTTPClient -MailboxName $Mailboxes[0]
		foreach ($MailboxName in $Mailboxes)
		{
			$rptObj = "" | Select-Object MailboxName, Language, Locale, TimeZone, AutomaticReplyStatus
			$EndPoint = Get-EXREndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings"
			$Results = Invoke-EXRRestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			$rptObj.MailboxName = $MailboxName
			$rptObj.Language = $Results.Language.DisplayName
			$rptObj.Locale = $Results.Language.Locale
			$rptObj.TimeZone = $Results.TimeZone
			$rptObj.AutomaticReplyStatus = $Results.AutomaticRepliesSetting.Status
			$rptCollection += $rptObj
		}
		Write-Output  $rptCollection
		
	}
}