function New-CalendarEvent {
	param(
		[Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
		[Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
		[Parameter(Position = 2, Mandatory = $true)] [psobject]$CalendarName,
		[Parameter(Position = 2, Mandatory = $true)] [psobject]$EventDetails
	)

	Begin {
		if ($AccessToken -eq $null) {
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
        
		$calendarID = Get-CalendarID -MailboxName $MailboxName -AccessToken $AccessToken -CalendarName $CalendarName

		if ($calendarID -eq $null) {
			# If the calendar doesn't exist yet, create it
			$calendarID = (New-CalendarFolder -MailboxName $user -AccessToken $Token -DisplayName $calName).ID
		}


		$RequestURL = $EndPoint + "('$MailboxName')/calendars/$calendarID/events"
		Write-Host "Request URL: $RequestURL"
    
		$JSONOutput = Invoke-RestPost -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $EventDetails
    
		$relevantData = $JSONOutput.value
		Write-Output $relevantData
	}
}