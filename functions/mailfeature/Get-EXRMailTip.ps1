function Get-EXRMailTips
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
        $Mailboxes,

        [Parameter(Position = 3, Mandatory = $false)]
		[String]
        $Tips

		
	)
	Begin
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
        $JsonPostDef = @{}
        $JsonPostDef.Add("EmailAddresses",$Mailboxes)
        $JsonPostDef.Add("MailTipsOptions",$Tips)
        $Content = ConvertTo-Json -InputObject $JsonPostDef -Depth 5
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users" -beta
		$RequestURL = $EndPoint + "/" + $MailboxName + "/GetMailTips"
		$Result = Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $Content
		return $Result.value
	}
}
