function New-EXRCalendarEvent {
	param(
		[Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
		[Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
		[Parameter(Position = 2, Mandatory = $true)] [psobject]$CalendarName,
		[Parameter(Position = 2, Mandatory = $true)] [psobject]$EventDetails
	)

	Begin {
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
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
        
		$calendarID = Get-EXRCalendarID -MailboxName $MailboxName -AccessToken $AccessToken -CalendarName $CalendarName

		if ($calendarID -eq $null) {
			# If the calendar doesn't exist yet, create it
			$calendarID = (New-EXRCalendarFolder -MailboxName $MailboxName -AccessToken $AccessToken -DisplayName $CalendarName).ID
		}


		$RequestURL = $EndPoint + "('$MailboxName')/calendars/$calendarID/events"
	   
		$JSONOutput = Invoke-RestPost -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $EventDetails
    
		$relevantData = $JSONOutput.value
		Write-Output $relevantData
	}
}