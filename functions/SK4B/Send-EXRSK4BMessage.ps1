function Send-EXRSK4BMessage{
    param(
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
         [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $AccessToken,
        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $Subject,
        [Parameter(Position = 4, Mandatory = $false)]
        [string]
        $ToSipAddress,
        [Parameter(Position = 5, Mandatory = $false)]
        [string]
        $Message

    )
    process{
      
 
        $MessageObject = @{}
        $MessageObject.Add("rel","service:startMessaging");
        if(![String]::IsNullOrEmpty($Subject)){
            $MessageObject.Add("subject",$Subject)
        }else{
            $MessageObject.Add("subject","")
        }
        $MessageObject.Add("operationId",[guid]::NewGuid().toString())
        if(![String]::IsNullOrEmpty($ToSipAddress)){
            $MessageObject.Add("to",("sip:" + $ToSipAddress))
        }else{
           throw ("Error you need to specify a repcipient")
        }
        if(![String]::IsNullOrEmpty($Message)){
            $MessageObject.Add("message",("data:text/plain," + $Message))
        }else{
           throw ("Error you need to specify a Message to send")
        }

        $HttpClient = Get-HTTPClient -MailboxName $Script:SK4BMailboxName
        $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $Script:SK4BToken.access_token));
        $URL =  ("https://" + $Script:SK4BServerName + $Script:SK4BApplication._embedded.communication._links.startMessaging.href) 
        $HttpClient.DefaultRequestHeaders.Add('X-MS-RequiresMinResourceVersion','2')
        $PostJson =  New-Object System.Net.Http.StringContent((ConvertTo-Json $MessageObject -Depth 9), [System.Text.Encoding]::UTF8, "application/json") 
        $ClientResult = $HttpClient.PostAsync([Uri]$URL,$PostJson)
        return ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
    }
}

