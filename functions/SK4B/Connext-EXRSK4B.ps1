function Connect-EXRSK4B {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,

        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $certificateFileName,

        [Parameter(Position = 2, Mandatory = $false)]
        [SecureString]
        $certificateFilePassword,

        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $ClientId
    )
    process {
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        }  
        $URL = 'https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root?originalDomain=' + $MailboxName.Split('@')[1]
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $ClientResult = $HttpClient.GetAsync([Uri]$URL)
        $QueryResult = ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
        if($QueryResult._links.redirect){            
            $HttpClient = Get-HTTPClient -MailboxName $MailboxName
            $ClientResult = $HttpClient.GetAsync([Uri]$QueryResult._links.redirect.href)
            $QueryResult = ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
        }
        $adURI = [URI]$QueryResult._links.user.href
        if ($certificateFileName) {
            $TenantId = Get-EXRTenantId -DomainName $MailboxName.Split('@')[1]
            $ucwatoken = Get-EXRAppOnlyToken -CertFileName $certificateFileName -TenantId $TenantId -ClientId $ClientId  -ResourceURL "NOAMmeetings.resources.lync.com" -MailboxName $MailboxName -password $certificateFilePassword
            $HttpClient = Get-HTTPClient -MailboxName $MailboxName
            $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $ucwatoken.access_token));
            $ClientResult = $HttpClient.GetAsync([Uri]"https://api.skypeforbusiness.com/platformservice/discover")
            $DiscoveryResult = ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result        
            $adURI2 = [URI]$DiscoveryResult._links.applications.href
            $ucwatoken = Get-EXRAppOnlyToken -CertFileName $certificateFileName -TenantId $TenantId -ClientId $ClientId  -ResourceURL $adURI2.host  -MailboxName $MailboxName -password $certificateFilePassword
        }
        else {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }         
            $ucwatoken = Invoke-RefreshAccessToken -AccessToken $AccessToken -MailboxName $MailboxName -ResourceURL $adURI.host -ucwa
            $HttpClient = Get-HTTPClient -MailboxName $MailboxName
            $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $ucwatoken.access_token));
            $ClientResult = $HttpClient.GetAsync([Uri]$QueryResult._links.user.href)
            $DiscoveryResult = ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result        
            $adURI2 = [URI]$DiscoveryResult._links.applications.href
            $ucwatoken = Invoke-RefreshAccessToken -AccessToken $AccessToken -MailboxName $MailboxName -ResourceURL $adURI2.host -ucwa
        }
        $EndPointId = [guid]::NewGuid().ToString();
        $EndPointPost = @'
         {
            "UserAgent":  "PSAgent",
            "Culture":  "en-US",
            "EndpointId":  "$EndPointId"
        }
'@
        $PostJson = New-Object System.Net.Http.StringContent($EndPointPost.Replace("`$EndPointId", $EndPointId), [System.Text.Encoding]::UTF8, "application/json") 
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $ucwatoken.access_token));
        $HttpClient.DefaultRequestHeaders.Add('X-MS-RequiresMinResourceVersion', '2')
        $ClientResult = $HttpClient.PostAsync([Uri]$DiscoveryResult._links.applications.href, $PostJson)        
        $EndPointResult = ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
        $Script:SK4BApplication = $EndPointResult
        $Script:SK4BServerName = $adURI2.host
        $Script:SK4BMailboxName = $MailboxName
        $Script:SK4BNextEvent = $EndPointResult._links.events.href       
        if ($EndPointResult._embedded.me._links.makeMeAvailable.href) {
            $MakeMeAvailblePost = @'
            {
                "SupportedModalities":  [
                                            "Messaging"
                                        ],
                "SupportedMessageFormats":  [
                                                "Plain",
                                                "Html"
                                            ]
            }

'@
            $HttpClient = Get-HTTPClient -MailboxName $MailboxName
            $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $ucwatoken.access_token));
            $URL = ("https://" + $adURI2.host + $EndPointResult._embedded.me._links.makeMeAvailable.href)
            $PostJson = New-Object System.Net.Http.StringContent($MakeMeAvailblePost, [System.Text.Encoding]::UTF8, "application/json") 
            $ClientResult = $HttpClient.PostAsync([Uri]$URL, $PostJson)            
        }
        $events = Invoke-EXRGetUCWAEvents -MailboxName $MailboxName
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $ucwatoken.access_token));
        $URL = ("https://" + $Script:SK4BServerName + $EndPointResult._embedded.me._links.self.href)
        $ClientResult = $HttpClient.GetAsync([Uri]$URL)   
        $Script:SK4BLinks = ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result   
        Write-Host ("Connected to sk4b for " + $Script:SK4BMailboxName)          
    }
}



