function Invoke-EXRRefreshAccessToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-EXRHTTPClient -MailboxName $MailboxName
		$ClientId = $AccessToken.clientid
		# $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AccessToken.redirectUrl)
		$redirectUrl = $AccessToken.redirectUrl
		$RefreshToken = (Get-EXRTokenFromSecureString -SecureToken $AccessToken.refresh_token)
		$AuthorizationPostRequest = "client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientResult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
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
			$JsonObject = ConvertFrom-Json -InputObject $ClientResult.Result.Content.ReadAsStringAsync().Result
			Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $AccessToken.clientid
			Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $AccessToken.redirectUrl
			if ([bool]($AccessToken.PSobject.Properties.name -match "Beta"))
			{
				Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
			}
			if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
			{
				$JsonObject.refresh_token = (Get-EXRProtectedToken -PlainToken $JsonObject.refresh_token)
			}
			if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
			{
				$JsonObject.access_token = (Get-EXRProtectedToken -PlainToken $JsonObject.access_token)
			}
			$HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
			if(!$MyInvocation.MyCommand.Module.PrivateData['EXRTokens'].ContainsKey($HostDomain)){			
				$MyInvocation.MyCommand.Module.PrivateData['EXRTokens'].Add($HostDomain,$JsonObject)
			}
			else{
				$MyInvocation.MyCommand.Module.PrivateData['EXRTokens'][$HostDomain] = $JsonObject
			}
			return $JsonObject
		}
		
	}
}
