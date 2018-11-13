function Get-EXRTenantId {
    param( 
        [Parameter(Position = 1, Mandatory = $false)]
        [String]$DomainName
       
    )  
    Begin {
        try{
            $RequestURL = "https://login.windows.net/{0}/.well-known/openid-configuration" -f $DomainName
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
            $ClientResult = $HttpClient.GetAsync([Uri]$RequestURL)
            $JsonResponse = ConvertFrom-Json  $ClientResult.Result.Content.ReadAsStringAsync().Result
            $ValArray = $JsonResponse.authorization_endpoint.replace("https://login.windows.net/","").split("/")
            return $ValArray[0]
        }catch{

        }

    }
}