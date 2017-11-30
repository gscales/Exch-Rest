

function  Send-EXRSimpleMeetingRequest {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$false)] [string]$SenderName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [String]$Attendee,
        [Parameter(Position=3, Mandatory=$false)] [DateTime]$Start,  
        [Parameter(Position=4, Mandatory=$false)] [DateTime]$End,
        [Parameter(Position=5, Mandatory=$false)] [String]$Subject        
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
        $Attendees = @()
        $Attendees += (new-attendee -Name $Attendee -Address $Attendee -type 'Required')
        New-EXRCalendarEventREST -MailboxName $SenderName -AccessToken $AccessToken -Start $Start -End $End -Subject $Subject -Attendees $Attendees
 
    }
}