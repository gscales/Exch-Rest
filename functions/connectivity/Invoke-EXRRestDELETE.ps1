function Invoke-EXRRestDELETE
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$RequestURL,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$MailboxName,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[System.Net.Http.HttpClient]
		$HttpClient,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[psobject]
		$AccessToken,
		[Parameter(Position=4, Mandatory=$false)] [switch]$NoJSON,
        [Parameter(Position=5, Mandatory=$false)] [bool]$TrackStatus = $false
		
	)
	Begin
	{
		#Check for expired Token
		$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
		$expiry = $minTime.AddSeconds($AccessToken.expires_on)
		if ($expiry -le [DateTime]::Now.ToUniversalTime())
		{
			write-host "Refresh Token"
			$AccessToken = Invoke-EXRRefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
			Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
		}
		$method = New-Object System.Net.Http.HttpMethod("DELETE")
		$HttpRequestMessage = New-Object System.Net.Http.HttpRequestMessage($method, [Uri]$RequestURL)
		$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-EXRTokenFromSecureString -SecureToken $AccessToken.access_token));
		$ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
             if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK){
                 if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created){
                     write-Host ($ClientResult.Result)
                 }
                 if($ClientResult.Result.Content -ne $null){
                    Write-Host ($ClientResult.Result.Content.ReadAsStringAsync().Result); 
                 }                  
             }             
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {
                    Write-Host ("Error making REST Get " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                    Write-Host ("RequestURL : " + $RequestURL)                
             }
            else
             {
               if($NoJSON){
                    return  $ClientResult.Result.Content  
               }
               else{
                    $JsonObject = ExpandPayload($ClientResult.Result.Content.ReadAsStringAsync().Result) 
                    #$JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
                   if([String]::IsNullOrEmpty($ClientResult)){
                        write-host "No Value returned"
                   }
                   else{
									
						if($ClientResult.Result.StatusCode -eq "NoContent"){
							Write-host "Item Deleted"
						}
						return $JsonObject	
                   }

               }  

             }
		
	}
}
