
function Get-EXRSK4BPresence{
    param(
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $TargetUser,
        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $PresenceURI,
        [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $AccessToken
    )
    process{
        
        $HttpClient = Get-HTTPClient -MailboxName $Script:SK4BMailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $Script:SK4BToken.access_token));
        if($TargetUser){
            $URL = ("https://" + $Script:SK4BServerName + $Script:SK4BLinks._links.self.href.replace("me","people") + "/" + $TargetUser + "/presence")
        }else{
            $URL = ("https://" + $Script:SK4BServerName + $PresenceURI)
        }        
        $ClientResult = $HttpClient.GetAsync([Uri]$URL)   
        return ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result   
         
    }
}