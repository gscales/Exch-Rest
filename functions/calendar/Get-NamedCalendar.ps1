function Get-NamedCalendarView {
	param( 
		[Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
		[Parameter(Position = 1, Mandatory = $true)] [string]$CalendarName,
		[Parameter(Position = 2, Mandatory = $false)] [psobject]$AccessToken,
		[Parameter(Position = 3, Mandatory = $true)] [psobject]$StartTime,
		[Parameter(Position = 4, Mandatory = $true)] [psobject]$EndTime
	)
	Begin {
		if ($AccessToken -eq $null) {
			$AccessToken = Get-AccessToken -MailboxName $MailboxName          
		}   
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$calendarID = Get-CalendarID -MailboxName $MailboxName -AccessToken $AccessToken -CalendarName $CalendarName
		$RequestURL = $EndPoint + "('$MailboxName')/calendars/$calendarID/calendarview?startdatetime=" + $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + "&enddatetime=" + $EndTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
		Write-Host $RequestURL
		do {
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value) {
				Write-Output $Message
			}           
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}while (![String]::IsNullOrEmpty($RequestURL))
	}
}