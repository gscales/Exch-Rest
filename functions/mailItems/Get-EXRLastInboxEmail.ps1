function Get-EXRLastInboxEmail{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Process{
	$Items = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder Inbox -TopOnly:$true -Top 1
	Get-EXREmail -ItemRESTURI $Items[0].ItemRESTURI -MailboxName $MailboxName -AccessToken $AccessToken
    }
}