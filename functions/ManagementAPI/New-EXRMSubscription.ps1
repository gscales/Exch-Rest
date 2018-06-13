function New-EXRMSubscription {
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
        $TenantId,

        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $ContentType,

        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $webhookAddress,

        [Parameter(Position = 4, Mandatory = $false)]
        [string]
        $webhookauthId,

        [Parameter(Position = 5, Mandatory = $false)]
        [DateTime]
        $webhookexpiration,
       
        [Parameter(Position = 6, Mandatory = $false)]
        [switch]
        $Exchange,
        
        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $AzureAD,

        [Parameter(Position = 8, Mandatory = $false)]
        [switch]
        $SharePoint,

        [Parameter(Position = 9, Mandatory = $false)]
        [switch]
        $DLP,

        [Parameter(Position = 10, Mandatory = $false)]
        [switch]
        $General

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
        if ([String]::IsNullOrEmpty($ContentType)) {
            if ($Exchange.IsPresent) {
                $ContentType = "Audit.Exchange"
            }      
            if ($AzureAD.IsPresent) {
                $ContentType = "Audit.AzureActiveDirectory"
            }   
            if ($SharePoint.IsPresent) {
                $ContentType = "Audit.SharePoint"
            }
            if ($DLP.IsPresent) {
                $ContentType = "DLP.All"
            }     
            if ($General.IsPresent) {
                $ContentType = "Audit.General"
            }  
        }
        $PublisherId = $TenantId
        $PostContent = ""
        if (![String]::IsNullOrEmpty($webhookAddress)) {
            $PostContent = "{`r`n `"webhook`" : {"
            $PostContent += "`"address`": `"" + $webhookAddress + "`""
            if (![String]::IsNullOrEmpty($webhookauthId)) {
                $PostContent += ",`r`n`"authId`": `"" + $webhookauthId + "`"" 
            }
            if (![String]::IsNullOrEmpty($webhookexpiration)) {
                $PostContent += ",`r`n`"expiration`": `"" + $webhookexpiration + "`"`r`n" 
            }       
            $PostContent += "}`r`n}"
        }
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $RequestURL = "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/start?contentType={1}&PublisherIdentifier={2}" -f $TenantId, $ContentType, $PublisherId 
        $JSONOutput = Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $PostContent
        return $JSONOutput 
		
    }
}
