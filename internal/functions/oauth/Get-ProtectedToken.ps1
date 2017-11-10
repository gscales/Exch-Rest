function Get-ProtectedToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$PlainToken
	)
	begin
	{
		$SecureEncryptedToken = Protect-String -String $PlainToken
		return, $SecureEncryptedToken
	}
}
