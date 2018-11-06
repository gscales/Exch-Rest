function Invoke-EXRSK4BReportMyActivity{
    param(
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
         [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $AccessToken
    )
    process{
        $PostContent = @{};
        $PostJson =  New-Object System.Net.Http.StringContent((ConvertTo-Json $PostContent -Depth 9), [System.Text.Encoding]::UTF8, "application/json") 
        $HttpClient = Get-HTTPClient -MailboxName $Script:SK4BMailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $Script:SK4BToken.access_token));
        $URL =  ("https://" + $Script:SK4BServerName + $Script:SK4BLinks._links.reportMyActivity.href)
        $ClientResult = $HttpClient.PostAsync([Uri]$URL,$PostJson)
        return ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
    }
    
}