function ConvertFrom-SecureStringCustom
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[System.Security.SecureString]
		$SecureToken
	)
	process
	{
		#$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
		$Token = Unprotect-String -String $SecureToken
		return, $Token
	}
}
