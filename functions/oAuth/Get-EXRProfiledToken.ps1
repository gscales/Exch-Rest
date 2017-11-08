function Get-EXRProfiledToken {      
    param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
        $MailboxName
    )
    Process{
        	$HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
			if($MyInvocation.MyCommand.Module.PrivateData['EXRTokens'].ContainsKey($HostDomain)){			
				return $MyInvocation.MyCommand.Module.PrivateData['EXRTokens'][$HostDomain]
			}

    }
}