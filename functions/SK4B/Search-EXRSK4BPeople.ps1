function Search-EXRSK4BPeople{
    param(
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
         [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $AccessToken,
        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $mail

    )
    process{
      

        $HttpClient = Get-HTTPClient -MailboxName $Script:SK4BMailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $Script:SK4BToken.access_token));
        $URL =  ("https://" + $Script:SK4BServerName + $Script:SK4BLinks._links.self.href.replace("me","people") + "/search?mail=" + [uri]::EscapeDataString($mail)) 
        $HttpClient.DefaultRequestHeaders.Add('X-MS-RequiresMinResourceVersion','2')
        $ClientResult = $HttpClient.GetAsync([Uri]$URL)
        return ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
    }
}