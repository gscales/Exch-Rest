function Get-EXRMSubscriptionContent {
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
        $ContentType,

        [Parameter(Position = 4, Mandatory = $false)]
        [DateTime]
        $StartTime,

        [Parameter(Position = 5, Mandatory = $false)]
        [DateTime]
        $EndTime,

        [Parameter(Position = 6, Mandatory = $false)]
        [switch]
        $returnContentBlobs,
        
        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $Exchange,
        
        [Parameter(Position = 8, Mandatory = $false)]
        [switch]
        $AzureAD,

        [Parameter(Position = 9, Mandatory = $false)]
        [switch]
        $SharePoint,

        [Parameter(Position = 10, Mandatory = $false)]
        [switch]
        $DLP,

        [Parameter(Position = 11, Mandatory = $false)]
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
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        if ([String]::IsNullOrEmpty($ContentType)) {throw "You must choose or entry a ContentType"}
        if ($StartTime -eq $null) {
            $RequestURL = "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/content?contentType={1}&PublisherIdentifier={2}" -f $TenantId, $ContentType, $PublisherId
        }
        else {
            $RequestURL = "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/content?contentType={1}&PublisherIdentifier={2}&startTime={3}&endTime={4}" -f $TenantId, $ContentType, $PublisherId, $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"), $EndTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        $NextPageHeader = ""
        do {
            $Output = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -RawResponse
            $RequestURL = ""
            $NextPageHeader = ""
            if ($Output -ne $null) {
                if ($Output.Result.Headers.Contains("NextPageUri")) {
                    foreach ($header in $Output.Result.Headers) {
                        if ($header.key -eq "NextPageUri") {
                            $NextPageHeader = $header.value[0]
                            $RequestURL = $NextPageHeader + "&PublisherIdentifier={0}" -f $PublisherId
                        }
                    }
                }
                $JSONOutput = ExpandPayload($Output.Result.Content.ReadAsStringAsync().Result)  
                if ($returnContentBlobs.IsPresent) {
                    foreach ($val in $JSONOutput) {
                        $blobRequest = $val.contentUri + "?PublisherIdentifier={0}" -f $PublisherId
                        $ContentBlob = Get-EXRMSubscriptionContentBlob -ContentURI $blobRequest
                        Add-Member -InputObject $val -NotePropertyName "ContentBlob" -NotePropertyValue $ContentBlob
                    }
                }
                Write-Output $JSONOutput
            }

        }while ($NextPageHeader -ne "")		 

         
		
    }
}
