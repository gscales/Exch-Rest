function Unprotect-String {
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
	
    begin {
        Add-Type -AssemblyName System.Security -ErrorAction Stop
    }
    process {
        if ($PSVersionTable.PSEdition -eq "Core") {        
            foreach ($item in $String) {
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($item)			
                $EncyptedToken = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                $DcyptedToken = ConvertTo-SecureString -String $EncyptedToken -Key $Script:EncKey
                $BSTR1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($DcyptedToken)
                [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR1)
            }
        }
        else {
            foreach ($item in $String) {
                $cred = New-Object PSCredential("irrelevant", $item)
                $stringBytes = [System.Convert]::FromBase64String($cred.GetNetworkCredential().Password)
                $decodedBytes = [System.Security.Cryptography.ProtectedData]::Unprotect($stringBytes, $null, 'CurrentUser')
                [Text.Encoding]::UTF8.GetString($decodedBytes)
            }
        }
    }
}