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
		$CacheCredentials,

		[Parameter(Position = 8, Mandatory = $false)]
		[string]
		$TenantId,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[switch]
		$DumpCache
		
	)
	Begin
	{
		#if($DumpCache.IsPresent){
	    #		write-output $Script:TokenCache
	    #	}
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
			if([String]::IsNullOrEmpty($TenantId)){
				$Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=$Prompt&response_mode=form_post"
			}else{
				$Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/$TenantId/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=$Prompt&response_mode=form_post"
			}
			
			$code = $Phase1auth
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
			Add-Member -InputObject $JsonObject -NotePropertyName mailbox -NotePropertyValue $MailboxName
			if(![String]::IsNullOrEmpty($TenantId)){
				Add-Member -InputObject $JsonObject -NotePropertyName TenantId -NotePropertyValue $TenantId
			}
			if ($Beta.IsPresent)
			{
				Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue $True
			}
			if($CacheCredentials.IsPresent){
				if(!$Script:TokenCache.ContainsKey($ResourceURL)){	
					$ResourceTokens = @{}		
					$Script:TokenCache.Add($ResourceURL,$ResourceTokens)
				}
				Add-Member -InputObject $JsonObject -NotePropertyName Cached -NotePropertyValue $true				
				$HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
				if(!$Script:TokenCache[$ResourceURL].ContainsKey($HostDomain)){			
					$Script:TokenCache[$ResourceURL].Add($HostDomain,$JsonObject)
				}
				else{
					$Script:TokenCache[$ResourceURL][$HostDomain] = $JsonObject
				}
				write-host ("Cached Token for " + $ResourceURL + " " + $HostDomain)
			}
			return $JsonObject
		}
	}
}
