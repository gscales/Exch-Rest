function Get-EXRMHistoricalStatus
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,

		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$TenantId

	)
	Process
	{
		$PublisherId = "5b24a168-aa6c-40db-b191-b509184797fb"
		if($AccessToken -eq $null)
        {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName -ResourceURL "manage.office.com" 
            if($AccessToken -eq $null){
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
         if([String]::IsNullOrEmpty($MailboxName)){
            $MailboxName = $AccessToken.mailbox
		} 
		if([String]::IsNullOrEmpty($TenantId)){
			$HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
			$TenantId = Get-EXRtenantId -Domain $HostDomain
		}
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$RequestURL = "https://manage.office.com/api/v1.0/{0}/ServiceComms/HistoricalStatus" -f $TenantId,$PublisherId 
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		return $JSONOutput.value 
		
	}
}
