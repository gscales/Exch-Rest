function Invoke-RestPut
{
	[CmdletBinding()]
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
		
		[Parameter(Position = 4, Mandatory = $true)]
		[String]
		$ContentHeader,
		
		[Parameter(Position = 5, Mandatory = $true)]
		[PSCustomObject]
		$Content
		
	)
	Begin
	{
		if($Script:TraceRequest){
			write-host $RequestURL
		}
		#Check for expired Token
		$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
		$expiry = $minTime.AddSeconds($AccessToken.expires_on)
		if ($expiry -le [DateTime]::Now.ToUniversalTime())
		{
			write-host "Refresh Token"
			$AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
		}
		$method = New-Object System.Net.Http.HttpMethod("PUT")
		$HttpRequestMessage = New-Object System.Net.Http.HttpRequestMessage($method, [Uri]$RequestURL)
		$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $AccessToken.access_token));
		if (![String]::IsNullorEmpty($ContentHeader))
		{
			$HttpRequestMessage.Content = New-Object System.Net.Http.ByteArrayContent -ArgumentList @( ,$Content)
		}
		else
		{
			$HttpRequestMessage.Content = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
		}
		$ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
		if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
		{
			if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created)
			{
				write-Output ($ClientResult.Result)
			}
			if ($ClientResult.Result.Content -ne $null)
			{
				Write-Output ($ClientResult.Result.Content.ReadAsStringAsync());
			}
		}
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Output ("Error making REST PUT " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			# $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
			$JsonObject = ExpandPayload -response $ClientResult.Result.Content.ReadAsStringAsync().Result
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
