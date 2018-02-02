function Get-EXRCalendarID {
	param(
		[Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
		[Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
		[Parameter(Position = 2, Mandatory = $true)] [psobject]$CalendarName
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
		$RequestURL = $EndPoint + "('$MailboxName')/calendars"
		do {
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value) {
				if ($message.Name -eq $CalendarName) {
					Write-Output $Message.Id
				}
			}      
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}while (![String]::IsNullOrEmpty($RequestURL))  
	}
}