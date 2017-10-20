function New-JWTToken
{
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$CertFileName,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$TenantId,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$ClientId,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[Int32]
		$ValidateForMinutes,
		
		[Parameter(Mandatory = $True)]
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
		$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertFileName, $password, $exVal
		$x5t = [System.Convert]::ToBase64String($cert.GetCertHash())
		$jti = [System.Guid]::NewGuid().ToString()
		$Headerassertaion = "{"
		$Headerassertaion += "     `"alg`": `"RS256`","
		$Headerassertaion += "     `"x5t`": `"" + $x5t + "`""
		$Headerassertaion += "}"
		$PayLoadassertaion += "{"
		$PayLoadassertaion += "    `"aud`": `"https://login.windows.net/" + $TenantId + "/oauth2/token`","
		$PayLoadassertaion += "    `"exp`": $exp,"
		$PayLoadassertaion += "    `"iss`": `"" + $ClientId + "`","
		$PayLoadassertaion += "    `"jti`": `"" + $jti + "`","
		$PayLoadassertaion += "    `"nbf`": $nbf,"
		$PayLoadassertaion += "    `"sub`": `"" + $ClientId + "`""
		$PayLoadassertaion += "} "
		$encodedHeader = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($Headerassertaion)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$encodedPayLoadassertaion = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($PayLoadassertaion)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$JWTOutput = $encodedHeader + "." + $encodedPayLoadassertaion
		$SigBytes = [System.Text.UTF8Encoding]::UTF8.GetBytes($JWTOutput)
		$rsa = $cert.PrivateKey;
		$sha256 = [System.Security.Cryptography.SHA256]::Create()
		$hash = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($encodedHeader + '.' + $encodedPayLoadassertaion));
		$sigform = New-Object System.Security.Cryptography.RSAPKCS1SignatureFormatter($rsa);
		$sigform.SetHashAlgorithm("SHA256");
		$sig = [System.Convert]::ToBase64String($sigform.CreateSignature($hash)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$JWTOutput = $encodedHeader + '.' + $encodedPayLoadassertaion + '.' + $sig
		Write-Output ($JWTOutput)
		
	}
}
