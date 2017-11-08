

function  Send-EXRSimpleMeetingRequest {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$SenderName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [String]$Attendee,
        [Parameter(Position=3, Mandatory=$false)] [DateTime]$Start,  
        [Parameter(Position=4, Mandatory=$false)] [DateTime]$End,
        [Parameter(Position=5, Mandatory=$false)] [String]$Subject        
    )
    Begin{
        $Attendees = @()
        $Attendees += (new-attendee -Name $Attendee -Address $Attendee -type 'Required')
        New-EXRCalendarEventREST -MailboxName $SenderName -AccessToken $AccessToken -Start $Start -End $End -Subject $Subject -Attendees $Attendees
 
    }
}