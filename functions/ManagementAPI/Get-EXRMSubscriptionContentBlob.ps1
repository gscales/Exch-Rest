function Get-EXRMSubscriptionContentBlob {
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
        $TenantId,

        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $ContentURI

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
        if(!$ContentURI.Contains("PublisherId")){
            $ContentURI = $ContentURI +  "?PublisherIdentifier={0}" -f  $TenantId
        }        
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $RequestURL = $ContentURI 
        $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        return $JSONOutput 
		
    }
}
