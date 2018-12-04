if($Script:TokenCache -eq $null){
	$Script:TokenCache = @{}
}
## Added to Support Powershell Core Encryption
$Script:EncKey = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Script:EncKey)