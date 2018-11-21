function Get-EXRMemberOf{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$UPN,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
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
    if([String]::IsNullOrEmpty($UPN)){
        $UPN = $MailboxName
    }
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment ""
        $RequestURL = $EndPoint + "me/memberOf"
        $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        foreach($Message in  $JSONOutput.Value){
            Write-Output $Message
        }
  
     
        
        
    }
}
