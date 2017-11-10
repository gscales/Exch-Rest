function Get-EXRAccessToken
{
	[CmdletBinding()]
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
		$Prompt,

		[Parameter(Position = 7, Mandatory = $false)]
		[switch]
		$SaveToPrivateData
		
	)
	Begin
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$AppSetting = Get-AppSettings
		if ([String]::IsNullOrEmpty($ClientId))
		{
			$ReturnToken = Get-ProfiledToken -MailboxName $MailboxName
			if($ReturnToken -eq $null){
				Write-Error ("No Access Token for " + $MailboxName)
			}
			else{
				return $ReturnToken
			}
		}
		else{
			if ([String]::IsNullOrEmpty(($ClientSecret)))
			{
				$ClientSecret = $AppSetting.ClientSecret
			}
			if ([String]::IsNullOrEmpty($redirectUrl))
			{
				$redirectUrl = [System.Web.HttpUtility]::UrlEncode("urn:ietf:wg:oauth:2.0:oob")
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
			if($SaveToPrivateData.IsPresent){
				$HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
				if(!$Script:TokenCache.ContainsKey($HostDomain)){			
					$Script:TokenCache.Add($HostDomain,$JsonObject)
				}
				else{
					$Script:TokenCache[$HostDomain] = $JsonObject
				}
			}
			return $JsonObject
		}
	}
}
