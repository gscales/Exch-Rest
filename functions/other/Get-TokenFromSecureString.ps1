function Get-TokenFromSecureString
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[System.Security.SecureString]
		$SecureToken
	)
	begin
	{
		$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
		$Token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
		return, $Token
	}
}
