function Get-AccessToken
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$ClientId,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$redirectUrl,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$ClientSecret,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$ResourceURL,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$Beta,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[String]
		$Prompt
		
	)
	Begin
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient($MailboxName)
		$AppSetting = Get-AppSettings
		if ($ClientId -eq $null)
		{
			$ClientId = $AppSetting.ClientId
		}
		if ($ClientSecret -eq $null)
		{
			$ClientSecret = $AppSetting.ClientSecret
		}
		if ($redirectUrl -eq $null)
		{
			$redirectUrl = [System.Web.HttpUtility]::UrlEncode($AppSetting.redirectUrl)
		}
		else
		{
			$redirectUrl = [System.Web.HttpUtility]::UrlEncode($redirectUrl)
		}
		if ([String]::IsNullOrEmpty($ResourceURL))
		{
			$ResourceURL = $AppSetting.ResourceURL
		}
		if ([String]::IsNullOrEmpty($Prompt))
		{
			$Prompt = "refresh_session"
		}
		
		$Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=$Prompt"
		$code = $Phase1auth["code"]
		$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
		if (![String]::IsNullOrEmpty($ClientSecret))
		{
			$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_secret=$ClientSecret&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
		}
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
		$JsonObject = ConvertFrom-Json -InputObject $ClientReesult.Result.Content.ReadAsStringAsync().Result
		if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
		{
			$JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
		{
			$JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "id_token"))
		{
			$JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
		}
		Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
		Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
		if ($Beta.IsPresent)
		{
			Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
		}
		return $JsonObject
	}
}
