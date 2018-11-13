function Get-EXRCredentialType {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName
		


    )
    Process {
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName       
        $RequestURL = "https://login.microsoftonline.com/?login_hint=" + $MailboxName
        Add-Type -AssemblyName System.Net.Http
        $handler = New-Object  System.Net.Http.HttpClientHandler
        $handler.CookieContainer = New-Object System.Net.CookieContainer
        $handler.AllowAutoRedirect = $true;
        $HttpClient = New-Object System.Net.Http.HttpClient($handler);
        $Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json")
        $HttpClient.DefaultRequestHeaders.Accept.Add($Header);
        $HttpClient.Timeout = New-Object System.TimeSpan(0, 0, 90);
        $HttpClient.DefaultRequestHeaders.TransferEncodingChunked = $false
        $Header = New-Object System.Net.Http.Headers.ProductInfoHeaderValue("RestClient", "1.1")
        $HttpClient.DefaultRequestHeaders.UserAgent.Add($Header);
        $Request = @{}
        #$Request.Add("username",$MailboxName)
        $Request.Add("Method","GetAuthMethods")
        $PostJson =  New-Object System.Net.Http.StringContent((ConvertTo-Json $Request -Depth 9), [System.Text.Encoding]::UTF8, "application/json")
        
        $ClientResult = $HttpClient.GetAsync([Uri]$RequestURL)
        $JsonResponse = ConvertTo-Json $ClientResult.Result.Content.ReadAsStringAsync().Result
        Write-Output $JsonResponse 
    }
}

