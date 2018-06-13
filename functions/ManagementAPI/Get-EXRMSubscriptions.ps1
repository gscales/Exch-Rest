function Get-EXRMSubscriptions {
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
    Process {
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName -ResourceURL "manage.office.com" 
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        } 
        if ([String]::IsNullOrEmpty($TenantId)) {
            $HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
            $TenantId = Get-EXRtenantId -Domain $HostDomain
        }
        $PublisherId = $TenantId
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $RequestURL = "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/list?PublisherIdentifier={1}" -f $TenantId, $PublisherId 
        $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        return $JSONOutput 
		
    }
}
