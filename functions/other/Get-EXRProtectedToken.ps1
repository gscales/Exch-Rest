function Get-EXRProtectedToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$PlainToken
	)
	begin
	{
		$SecureString = New-Object System.Security.SecureString
		for ($i = 0; $i -lt $PlainToken.length; $i++)
		{
			$SecureString.AppendChar($PlainToken[$i])
		}
		$EncryptedToken = ConvertFrom-SecureString -SecureString $SecureString
		$SecureEncryptedToken = ConvertTo-SecureString -String $EncryptedToken
		return, $SecureEncryptedToken
	}
}
