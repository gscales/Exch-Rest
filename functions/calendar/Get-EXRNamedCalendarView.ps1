function Get-EXRNamedCalendarView {
	param( 
		[Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
		[Parameter(Position = 1, Mandatory = $true)] [string]$CalendarName,
		[Parameter(Position = 2, Mandatory = $false)] [psobject]$AccessToken,
		[Parameter(Position = 3, Mandatory = $true)] [psobject]$StartTime,
		[Parameter(Position = 4, Mandatory = $true)] [psobject]$EndTime
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
		$RequestURL = $EndPoint + "('$MailboxName')/calendars/$calendarID/calendarview?startdatetime=" + $StartTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + "&enddatetime=" + $EndTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
		do {
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value) {
				Write-Output $Message
			}           
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}while (![String]::IsNullOrEmpty($RequestURL))
	}
}