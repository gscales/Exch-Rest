function Get-EXRAccessTokenUserAndPass
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
		$ResourceURL,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[switch]
		$Beta
	)
	Begin
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$AppSetting = Get-AppSettings
		if ($ClientId -eq $null)
		{
			$ClientId = $AppSetting.ClientId
		}
		
		if ([String]::IsNullOrEmpty($ResourceURL))
		{
			$ResourceURL = $AppSetting.ResourceURL
		}
		$UserName = $Credentials.UserName.ToString()
		$password = $Credentials.GetNetworkCredential().password.ToString()
		$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=password&username=$username&password=$password"
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
		$JsonObject = ConvertFrom-Json -InputObject $ClientReesult.Result.Content.ReadAsStringAsync().Result
		Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
		if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
		{
			$JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
		{
			$JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
		}
		if ($Beta.IsPresent)
		{
			Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "id_token"))
		{
			$JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
		}
		#Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
		return $JsonObject
	}
}
