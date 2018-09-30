function Set-EXRCalendarEvent {
	param(
		[Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
		[Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
		[Parameter(Position = 2, Mandatory = $true)] [string]$EventId,
		[Parameter(Position = 3, Mandatory = $false)] [string]$Subject,
		[Parameter(Position = 4, Mandatory = $false)] [DateTime]$Start,
		[Parameter(Position = 5, Mandatory = $false)] [DateTime]$End,
		[Parameter(Position = 6, Mandatory = $false)] [String]$TimeZone,
		[Parameter(Position = 7, Mandatory = $false)]
		[ValidateSet('free','tentative','busy','oof','workingElsewhere','unknown')]
		[string]$ShowAs,
		[Parameter(Position = 8, Mandatory = $false)]
		[ValidateSet('normal','personal','private','confidential')]
		[string]$Sensitivity,
		[Parameter(Position = 9, Mandatory = $false)]
		[int32]$reminderMinutesBeforeStart,
		[Parameter(Position = 10, Mandatory = $false)]
		[bool]$isReminderOn,
		[Parameter(Position = 11, Mandatory = $false)]
		[string]$categories,
		[Parameter(Position = 12, Mandatory = $false)]
		[ValidateSet('low','normal','high')]
		[string]$importance,
		[Parameter(Position = 13, Mandatory = $false)] [switch] $ShowRequest
	)

	Begin {
		if($AccessToken -eq $null)
		{
			$AccessToken = Get-ProfiledToken -MailboxName $MailboxName
			if($AccessToken -eq $null)
			{
				$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
			}
		}
		if([String]::IsNullOrEmpty($MailboxName)){
			$MailboxName = $AccessToken.mailbox
		}
		
		if ([String]::IsNullOrEmpty($TimeZone))
		{
			$TimeZone = [TimeZoneInfo]::Local.Id
		}
		
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"

		$RequestURL = $EndPoint + "('$MailboxName')/calendar/events/$EventId"

		$NewEventDetails = "{" + "`r`n"
		
		if (![String]::IsNullOrEmpty($Subject))
		{
			$NewEventDetails += "`"Subject`": `"" + $Subject + "`"" + ",`r`n"
		}
		
		if ($Start -ne $null)
		{
			$NewEventDetails += "`"Start`": {   `"DateTime`":`"" + $Start.ToString("yyyy-MM-ddTHH:mm:ss") + "`"," + "`r`n"
			$NewEventDetails += "  `"TimeZone`":`"" + $TimeZone + "`"}" + ",`r`n"
		}
		if ($End -ne $null)
		{
			$NewEventDetails += "`"End`": {   `"DateTime`":`"" + $End.ToString("yyyy-MM-ddTHH:mm:ss") + "`"," + "`r`n"
			$NewEventDetails += "  `"TimeZone`":`"" + $TimeZone + "`"}" + ",`r`n"
		}
		
		if (![String]::IsNullOrEmpty($ShowAs))
		{
			$NewEventDetails += "`"ShowAs`": `"" + $ShowAs + "`"" + ",`r`n"
		}
		
		if (![String]::IsNullOrEmpty($Sensitivity))
		{
			$NewEventDetails += "`"Sensitivity`": `"" + $Sensitivity + "`"" + ",`r`n"
		}
		
		if (![String]::IsNullOrEmpty($categories))
		{
			$NewEventDetails += "`"categories`": `"" + $categories + "`"" + ",`r`n"
		}
		if (![String]::IsNullOrEmpty($importance))
		{
			$NewEventDetails += "`"importance`": `"" + $importance + "`"" + ",`r`n"
		}
		
		if ($PSBoundParameters.ContainsKey('reminderMinutesBeforeStart'))
		{
			$NewEventDetails += "`"reminderMinutesBeforeStart`": `"" + $reminderMinutesBeforeStart + "`"" + ",`r`n"
		}
		
		if ($PSBoundParameters.ContainsKey('isReminderOn'))
		{
			if ($isReminderOn)
			{
				$NewEventDetails += "`"isReminderOn`": true,`r`n"
			}
			else
			{
				$NewEventDetails += "`"isReminderOn`": false,`r`n"
			}
		}

		$NewEventDetails += "}"
		
		if ($ShowRequest.IsPresent)
		{
			Write-Host $NewEventDetails
		}
		
		$JSONOutput = Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewEventDetails

		$relevantData = $JSONOutput.value
		Write-Output $relevantData
	}
}
