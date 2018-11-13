function New-EXRJWTToken
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$CertFileName,

		[Parameter(Position = 2, Mandatory = $false)]
		[System.Security.Cryptography.X509Certificates.X509Certificate2]
		$Certificate,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$TenantId,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[string]
		$ClientId,
		
		[Parameter(Position = 5, Mandatory = $true)]
		[Int32]
		$ValidateForMinutes,
		
		[Parameter(Mandatory = $false)]
		[Security.SecureString]
		$password
	)
	Begin
	{
		
		$date1 = Get-Date -Date "01/01/1970"
		$date2 = (Get-Date).ToUniversalTime().AddMinutes($ValidateForMinutes)
		$date3 = (Get-Date).ToUniversalTime().AddMinutes(-5)
		$exp = [Math]::Round((New-TimeSpan -Start $date1 -End $date2).TotalSeconds, 0)
		$nbf = [Math]::Round((New-TimeSpan -Start $date1 -End $date3).TotalSeconds, 0)
		$exVal = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
		if(![String]::IsNullOrEmpty($CertFileName)){
			$Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertFileName, $password, $exVal
		}		
		$x5t = [System.Convert]::ToBase64String($Certificate.GetCertHash())
		$jti = [System.Guid]::NewGuid().ToString()
		$headerAssertion = @"
{
     "alg": "RS256",
     "x5t": "$x5t"
}
"@
		$payLoadAssertion += @"
{
    "aud": "https://login.windows.net/$TenantId/oauth2/token",
    "exp": $exp,
    "iss": "$ClientId",
    "jti": "$jti",
    "nbf": $nbf,
    "sub": "$ClientId"
}
"@
		$encodedHeader = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($headerAssertion)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$encodedPayLoadAssertion = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($payLoadAssertion)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$JWTOutput = $encodedHeader + "." + $encodedPayLoadAssertion
		$SigBytes = [System.Text.UTF8Encoding]::UTF8.GetBytes($JWTOutput)
		$rsa = $Certificate.PrivateKey;
		$sha256 = [System.Security.Cryptography.SHA256]::Create()
		$hash = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($encodedHeader + '.' + $encodedPayLoadAssertion));
		$sigform = New-Object System.Security.Cryptography.RSAPKCS1SignatureFormatter($rsa);
		$sigform.SetHashAlgorithm("SHA256");
		$sig = [System.Convert]::ToBase64String($sigform.CreateSignature($hash)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$JWTOutput = $encodedHeader + '.' + $encodedPayLoadAssertion + '.' + $sig
		Write-Output ($JWTOutput)
		
	}
}
