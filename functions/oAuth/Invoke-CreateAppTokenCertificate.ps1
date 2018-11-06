function Invoke-EXRCreateAppTokenCertificate {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $CertName,
		
        [Parameter(Position = 1, Mandatory = $true)]
        [string]
        $CertFileName,
        
        [Parameter(Position = 2, Mandatory = $true)]
        [string]
        $ObjectId
    )
    Begin {
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        }  
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName

        $Cert = New-SelfSignedCertificate -certstorelocation cert:\currentuser\my -dnsname $CertName -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'
        $SecurePassword = Read-Host -Prompt "Enter password for Certificate File" -AsSecureString
        $CertPath = "cert:\currentuser\my\" + $Cert.Thumbprint.ToString()
        Export-PfxCertificate -cert $CertPath -FilePath $CertFileName -Password $SecurePassword
        $bin = $cert.RawData
        $base64Value = [System.Convert]::ToBase64String($bin)
        $bin = $cert.GetCertHash()
        $base64Thumbprint = [System.Convert]::ToBase64String($bin)
        $keyid = [System.Guid]::NewGuid().ToString()
        Remove-Item $CertPath
        $RequestURL = "https://graph.microsoft.com/beta/applications('" + $ObjectId + "')"
        $PostContent = @{}
        $PostContent.Add("keyCredentials", @(@{ customKeyIdentifier = $base64Thumbprint; keyId = $keyid; type = "AsymmetricX509Cert"; usage = "Verify"; key = $base64Value }))
        $JsonPost = ConvertTo-Json -Depth 10 -InputObject $PostContent
		$JSONOutput = Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $JsonPost
		if($JSONOutput.IsSuccessStatusCode){
			Return "Successfully created"
		}else{
			Return $JSONOutput
		}
        
		
    }
	
}