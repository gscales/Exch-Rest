function Get-AppSettings
{
	[CmdletBinding()]
	param (
		
	)
	Begin
	{
		$configObj = "" | Select-Object ResourceURL, ClientId, redirectUrl, ClientSecret, x5t, TenantId, ValidateForMinutes
		$configObj.ResourceURL = "outlook.office.com"
		$configObj.ClientId = "" # 1bdbfb41-f690-4f93-b0bb-002004bbca79
		$configObj.redirectUrl = "" # http://localhost:8000/authorize
		$configObj.TenantId = "" # 1c3a18bf-da31-4f6c-a404-2c06c9cf5ae4
		$configObj.ClientSecret = ""
		$configObj.x5t = "" # VS/H6cNa/3gc9FrSxGs9jOOZP3o=
		$configObj.ValidateForMinutes = 60
		return $configObj
	}
}
