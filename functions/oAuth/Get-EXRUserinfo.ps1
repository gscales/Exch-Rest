function Get-EXRUserInfo
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$TargetUser
		
	)
	Begin
	{
		
		if($AccessToken -eq $null)
        {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if($AccessToken -eq $null){
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
         if([String]::IsNullOrEmpty($MailboxName)){
            $MailboxName = $AccessToken.mailbox
        } 
        $RequestURL = "https://login.windows.net/{0}/.well-known/openid-configuration" -f $MailboxName.Split('@')[1]
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
        $Result = Invoke-RestGet -RequestURL $JsonResponse.userinfo_endpoint -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -IdToken
		return $Result
	}
}
