function Invoke-EXRGetUCWAEvents{
    param(
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
        [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $AccessToken
    )
    process{
        Write-Verbose $Script:SK4BNextEvent          
        $HttpClient = Get-HTTPClient -MailboxName $Script:SK4BMailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $Script:SK4BToken.access_token));
        $URL =  ("https://" + $Script:SK4BServerName + $Script:SK4BNextEvent  + "&timeout=60")
        $ClientResult = $HttpClient.GetAsync([Uri]$URL)
        $events = ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
        $Script:SK4BNextEvent  = $events._links.next.href
	return $events
       
    } 
}
