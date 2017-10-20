function Invoke-RestDELETE
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
		$AccessToken
		
	)
	Begin
	{
		#Check for expired Token
		$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
		$expiry = $minTime.AddSeconds($AccessToken.expires_on)
		if ($expiry -le [DateTime]::Now.ToUniversalTime())
		{
			write-host "Refresh Token"
			$AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
			Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
		}
		$method = New-Object System.Net.Http.HttpMethod("DELETE")
		$HttpRequestMessage = New-Object System.Net.Http.HttpRequestMessage($method, [Uri]$RequestURL)
		$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
		$ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
		if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
		{
			if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::NoContent)
			{
				write-Host ($ClientResult.Result)
			}
			if ($ClientResult.Result.Content -ne $null)
			{
				Write-Output ($ClientResult.Result.Content.ReadAsStringAsync());
			}
		}
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Output ("Error making REST Delete " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			$JsonObject = ConvertFrom-Json -InputObject $ClientResult.Result.Content.ReadAsStringAsync().Result
			if ([String]::IsNullOrEmpty($JsonObject))
			{
				Write-Output $ClientResult.Result
			}
			else
			{
				return $JsonObject
			}
			
		}
		
	}
}
