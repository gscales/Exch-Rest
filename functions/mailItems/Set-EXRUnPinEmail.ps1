function Set-EXRUnPinEmail
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,

		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$ItemRESTURI
	)
	Process
	{
		if($AccessToken -eq $null)
        {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if($AccessToken -eq $null){
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
         if([String]::IsNullOrEmpty($MailboxName)){
            $MailboxName = $AccessToken.mailbox
        } 
		$Props = Get-EXRPinnedEmailProperty
		$Props[0].Value = "null"
		$Props[1].Value = "null"
		return Update-EXRMessage -MailboxName $MailboxName -ItemURI $ItemRESTURI -ExPropList $Props -AccessToken $AccessToken
		
	}
}
