function Get-CalendarID {
	param(
		[Parameter(Position = 0, Mandatory = $true)] [string]$MailboxName,
		[Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
		[Parameter(Position = 2, Mandatory = $true)] [psobject]$CalendarName
	)
	Begin {
		if ($AccessToken -eq $null) {
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/calendars"
		Write-Host $RequestURL
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