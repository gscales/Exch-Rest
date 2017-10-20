function Invoke-DecodeToken
{
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Token
	)
	## Start Code Attribution
	## Decode-Token function is based on work of the following Authors and should remain with the function if copied into other scripts
	## https://gallery.technet.microsoft.com/JWT-Token-Decode-637cf001
	## End Code Attribution
	Begin
	{
		$parts = $Token.Split('.');
		$headers = [System.Text.Encoding]::UTF8.GetString((Convert-FromBase64StringWithNoPadding $parts[0]))
		$claims = [System.Text.Encoding]::UTF8.GetString((Convert-FromBase64StringWithNoPadding $parts[1]))
		$signature = (Convert-FromBase64StringWithNoPadding $parts[2])
		
		$customObject = [PSCustomObject]@{
			headers  = ($headers | ConvertFrom-Json)
			claims   = ($claims | ConvertFrom-Json)
			signature = $signature
		}
		return $customObject
	}
}
