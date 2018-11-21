function Get-EXRGroupCalendar {
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
        $GroupId,

		[Parameter(Position = 3, Mandatory = $false)] [psobject]$StartTime=(Get-Date).Date,
        [Parameter(Position = 4, Mandatory = $false)] [psobject]$EndTime=(Get-Date).AddDays(7),
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Export
        


    )
    Process {

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
        $PropList = Get-EXRKnownProperty -PropertyName "AppointmentDuration"      

        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "groups" 
        $RequestURL = $EndPoint + "/" + $GroupId + "/calendarView?startdatetime=" + $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + "&enddatetime=" + $EndTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") 
        $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
        $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        foreach($Message in $JSONOutput.value){
            Expand-ExtendedProperties -Item $Message
            if($Export.IsPresent){
                $rptObj = "" | Select StartTime,EndTime,Duration,Type,Subject,Location,Organizer,Attendees,Notes,HasAttachments,IsReminderSet
                $rptObj.StartTime = ([DateTime]$Message.Start.dateTime).ToString("yyyy-MM-dd HH:mm")  
                $rptObj.EndTime = ([DateTime]$Message.End.dateTime).ToString("yyyy-MM-dd HH:mm")  
                $rptObj.Duration = $Message.AppointmentDuration
                $rptObj.Subject  = $Message.Subject   
                $rptObj.Type = $Message.type
                $rptObj.Location = $Message.Location.displayName
                $rptObj.Organizer = $Message.organizer.emailAddress.address
                $rptObj.HasAttachments = $Message.hasAttachments
                $rptObj.IsReminderSet = $Message.IsReminderSet
                foreach($attendee in $Message.attendees){
                    $atn = $attendee.emailaddress.address + " " + $attendee.type + ";"
                    $rptObj.Attendees += $atn
                }
                $rptObj.Notes = $Message.Body.content
                Write-Output $rptObj
            
            }else{
                Write-Output $Message
            }
           
        }
        
	
    }
}
