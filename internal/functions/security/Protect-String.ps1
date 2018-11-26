function Protect-String
{
<#
	.SYNOPSIS
		Uses DPAPI to encrypt strings.
	
	.DESCRIPTION
		Uses DPAPI to encrypt strings.
	
	.PARAMETER String
		The string to encrypt.
	
	.EXAMPLE
		PS C:\> Protect-String -String $secret
	
		Encrypts the content stored in $secret and returns it.
#>
	[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "")]
	[CmdletBinding()]
	Param (
		[Parameter(ValueFromPipeline = $true)]
		[string[]]
		$String
	)
	
	begin
	{
		Add-Type -AssemblyName System.Security -ErrorAction Stop
	}
	process
	{
		if($PSVersion.PSEdition -eq "Core"){
			$SecureString = New-Object System.Security.SecureString	 	
			for ($i = 0; $i -lt $String.length; $i++)	 		
			{	 		
				$SecureString.AppendChar($PlainToken[$i])	 			
			}	 		
			$EncryptedToken = ConvertFrom-SecureString -SecureString $SecureString	 		
			$SecureEncryptedToken = ConvertTo-SecureString -String $EncryptedToken
			return $SecureEncryptedToken
		}else{
			foreach ($item in $String)
			{
				$stringBytes = [Text.Encoding]::UTF8.GetBytes($item)
				$encodedBytes = [System.Security.Cryptography.ProtectedData]::Protect($stringBytes, $null, 'CurrentUser')
				[System.Convert]::ToBase64String($encodedBytes) | ConvertTo-SecureString -AsPlainText -Force
			}
		}
	}
}