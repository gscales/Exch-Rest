function Invoke-EXRCreateSelfSignedCert
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$CertName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$CertFileName,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$KeyFileName
	)
	Begin
	{
		$Cert = New-SelfSignedCertificate -certstorelocation cert:\currentuser\my -dnsname $CertName -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'
		$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
		$CertPath = "cert:\currentuser\my\" + $Cert.Thumbprint.ToString()
		Export-PfxCertificate -cert $CertPath -FilePath $CertFileName -Password $SecurePassword
		$bin = $cert.RawData
		$base64Value = [System.Convert]::ToBase64String($bin)
		$bin = $cert.GetCertHash()
		$base64Thumbprint = [System.Convert]::ToBase64String($bin)
		$keyid = [System.Guid]::NewGuid().ToString()
		$jsonObj = @{ customKeyIdentifier = $base64Thumbprint; keyId = $keyid; type = "AsymmetricX509Cert"; usage = "Verify"; value = $base64Value }
		$keyCredentials = ConvertTo-Json @($jsonObj) | Out-File $KeyFileName
		Remove-Item $CertPath
		Write-Host ("Key written to " + $KeyFileName)
		
	}
	
}
