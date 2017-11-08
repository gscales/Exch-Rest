function Invoke-EXRRestPOST
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
		[PSCustomObject]
		$Content
	)
	Begin
	{
		#Check for expired Token
		$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
		$expiry = $minTime.AddSeconds($AccessToken.expires_on)
		if ($expiry -le [DateTime]::Now.ToUniversalTime())
		{
			if ([bool]($AccessToken.PSobject.Properties.name -match "refresh_token"))
			{
				write-host "Refresh Token"
				$AccessToken = Invoke-EXRRefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
				Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
			}
			else
			{
				throw "App Token has expired a new access token is required rerun get-apptoken"
			}
		}
		$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-EXRTokenFromSecureString -SecureToken $AccessToken.access_token));
		$PostContent = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
		$HttpClient.DefaultRequestHeaders.Add("Prefer", ("outlook.timezone=`"" + [TimeZoneInfo]::Local.Id + "`""))
		$ClientResult = $HttpClient.PostAsync([Uri]($RequestURL), $PostContent)
		if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
		{
			if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created)
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
			Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			$JsonObject = ExpandPayload($ClientResult.Result.Content.ReadAsStringAsync().Result)
			#$JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
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
