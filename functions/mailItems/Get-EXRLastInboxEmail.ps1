function Get-EXRLastInboxEmail{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Process{
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
	$Items = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder Inbox -TopOnly:$true -Top 1
	Get-EXREmail -ItemRESTURI $Items[0].ItemRESTURI -MailboxName $MailboxName -AccessToken $AccessToken
    }
}