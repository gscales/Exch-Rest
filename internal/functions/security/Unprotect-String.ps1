function Unprotect-String
{
<#
	.SYNOPSIS
		Uses DPAPI to decrypt strings.
	
	.DESCRIPTION
		Uses DPAPI to decrypt strings.
		Designed to reverse encryption applied by Protect-String
	
	.PARAMETER String
		The string to decrypt.
	
	.EXAMPLE
		PS C:\> Unprotect-String -String $secret
	
		Decrypts the content stored in $secret and returns it.
#>
	[CmdletBinding()]
	Param (
		[Parameter(ValueFromPipeline = $true)]
		[System.Security.SecureString[]]
		$String
	)
	
	begin
	{
		Add-Type -AssemblyName System.Security -ErrorAction Stop
	}
	process
	{
		foreach ($item in $String)
		{
			$cred = New-Object PSCredential("irrelevant", $item)
			$stringBytes = [System.Convert]::FromBase64String($cred.GetNetworkCredential().Password)
			$decodedBytes = [System.Security.Cryptography.ProtectedData]::Unprotect($stringBytes, $null, 'CurrentUser')
			[Text.Encoding]::UTF8.GetString($decodedBytes)
		}
	}
}