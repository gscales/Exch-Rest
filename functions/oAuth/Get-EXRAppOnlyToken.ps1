function Get-EXRAppOnlyToken
{
	[CmdletBinding()]
	param (
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$CertFileName,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$TenantId,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$ClientId,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$redirectUrl,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[Int32]
		$ValidateForMinutes,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[string]
		$ResourceURL,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[switch]
		$Beta,
		
		[Parameter(Mandatory = $false)]
		[Security.SecureString]
		$password, 

		[Parameter(Position = 10, Mandatory = $false)]
		[string]
		$MailboxName,

		[Parameter(Position = 11, Mandatory = $false)]
		[switch]
		$NoCache,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[System.Security.Cryptography.X509Certificates.X509Certificate2]
		$Certificate
		
	)
	Begin
	{
		$AppSetting = Get-AppSettings
		if ($TenantId -eq $null)
		{
			$AppSetting.TenantId
		}
		if ($ClientId -eq $null)
		{
			$ClientId = $AppSetting.ClientId
		}
		if ($redirectUrl -eq $null)
		{
			$redirectUrl = $AppSetting.redirectUrl
		}
		if ($ValidateForMinutes -eq 0)
		{
			$ValidateForMinutes = $AppSetting.ValidateForMinutes
		}
		if ([String]::IsNullOrEmpty($ResourceURL))
		{
			$ResourceURL = $AppSetting.ResourceURL
		}
		if(![String]::IsNullOrEmpty($CertFileName)){
			$JWTToken = New-EXRJWTToken -CertFileName $CertFileName -password $password -TenantId $TenantId -ClientId $ClientId -ValidateForMinutes $ValidateForMinutes
		}else{
			$JWTToken = New-EXRJWTToken -Certificate $Certificate -TenantId $TenantId -ClientId $ClientId -ValidateForMinutes $ValidateForMinutes
		}		
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient -MailboxName " "
		$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_assertion_type=urn%3Aietf%3Aparams%3Aoauth%3Aclient-assertion-type%3Ajwt-bearer&client_assertion=$JWTToken&grant_type=client_credentials&redirect_uri=$redirectUrl"
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/" + $TenantId + "/oauth2/token"), $content)
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
		Add-Member -InputObject $JsonObject -NotePropertyName tenantid -NotePropertyValue $TenantId
		Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
        Add-Member -InputObject $JsonObject -NotePropertyName mailbox -NotePropertyValue $MailboxName
		Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
		if ($Beta.IsPresent)
		{
			Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
		}
		if(!$NoCache.IsPresent)		
		{
			if($MailboxName){
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
		}

		return $JsonObject
	}
}
