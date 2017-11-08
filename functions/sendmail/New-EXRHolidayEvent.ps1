function New-EXRHolidayEvent
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$CalendarName,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[PSCustomObject]
		$Calendar,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[String]
		$Subject,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 6, Mandatory = $true)]
		[datetime]
		$Day,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$SenderEmailAddress,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[psobject]
		$ReferanceAttachments,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[psobject]
		$Attendees,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[psobject]
		$ExPropList,
		
		[Parameter(Position = 14, Mandatory = $false)]
		[psobject]
		$StandardPropList,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[switch]
		$ShowRequest,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[switch]
		$RequestReadRecipient,
		
		[Parameter(Position = 19, Mandatory = $false)]
		[switch]
		$RequestDeliveryRecipient,
		
		[Parameter(Position = 20, Mandatory = $false)]
		[psobject]
		$ReplyTo,
		
		[Parameter(Position = 21, Mandatory = $false)]
		[string]
		$TimeZone,
		
		[Parameter(Position = 22, Mandatory = $false)]
		[psobject]
		$Recurrence,
		
		[Parameter(Position = 23, Mandatory = $false)]
		[switch]
		$Group,
		
		[Parameter(Position = 24, Mandatory = $false)]
		[string]
		$GroupName
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		if ([String]::IsNullOrEmpty($TimeZone))
		{
			$TimeZone = [TimeZoneInfo]::Local.Id
		}
		if ($AccessToken.resource -eq "https://graph.microsoft.com")
		{
			$isAllDay = Get-EXRItemProp -Name isAllDay -NoQuotes -Value true
		}
		else
		{
			$isAllDay = Get-EXRItemProp -Name IsAllDay -NoQuotes -Value true
		}
		if ($Calendar -eq $null)
		{
			$Calendar = Get-EXRDefaultCalendarFolder -MailboxName $MailboxName -AccessToken $AccessToken
		}
		if (![String]::IsNullOrEmpty($CalendarName))
		{
			$Calendar = Get-EXRCalendarFolder -MailboxName $MailboxName -AccessToken $AccessToken -FolderName $CalendarName
			if ([String]::IsNullOrEmpty($Calendar)) { throw "Error Calendar folder not found check the folder name this is case sensitive" }
		}
		if ($StandardPropList -eq $null)
		{
			$StandardPropList += $isAllDay
		}
		else
		{
			$StandardPropList = @()
			$StandardPropList += $isAllDay
		}
		$NewMessage = Get-EXREventJSONFormat -Subject $Subject -Body $Body -SenderEmailAddress $SenderEmailAddress -Start $Day.Date -End $Day.Date.AddDays(1) -TimeZone $TimeZone -Attachments $Attachments -ReferanceAttachments $ReferanceAttachments -Attendees $Attendees -SentDate $SentDate -ExPropList $ExPropList -StandardPropList $StandardPropList -ReplyTo $ReplyTo -RequestReadRecipient $RequestReadRecipient.IsPresent -RequestDeliveryRecipient $RequestDeliveryRecipient.IsPresent -Recurrence $Recurrence
		if ($ShowRequest.IsPresent)
		{
			write-host $NewMessage
		}
		
		if ($Group.IsPresent)
		{
			$EndPoint = Get-EXREndPoint -AccessToken $AccessToken -Segment "groups"
			$ModernGroup = Get-EXRModernGroups -MailboxName $MailboxName -GroupName $GroupName -AccessToken $AccessToken
			$RequestURL = $EndPoint + "('" + $ModernGroup.id + "')/events"
		}
		else
		{
			$EndPoint = Get-EXREndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "/" + $MailboxName + "/calendars('" + $Calendar.id + "')/events"
		}
		$HttpClient = Get-EXRHTTPClient -MailboxName $MailboxName
		return Invoke-EXRRestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage
		
	}
}
