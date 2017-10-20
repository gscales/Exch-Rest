#region Sending_Email
function New-SentEmailMessage
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[PSCustomObject]
		$Folder,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[String]
		$Subject,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 7, Mandatory = $true)]
		[psobject]
		$SenderEmailAddress,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[psobject]
		$ToRecipients,
		
		[Parameter(Position = 10, Mandatory = $true)]
		[DateTime]
		$SentDate,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[psobject]
		$ExPropList,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[string]
		$ItemClass
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$SentFlag = Get-TaggedProperty -DataType "Integer" -Id "0x0E07" -Value "1"
		$SentTime = Get-TaggedProperty -DataType "SystemTime" -Id "0x0039" -Value $SentDate.ToString("yyyy-MM-ddTHH:mm:ss.ffffzzz")
		$RcvdTime = Get-TaggedProperty -DataType "SystemTime" -Id "0x0E06" -Value $SentDate.ToString("yyyy-MM-ddTHH:mm:ss.ffffzzz")
		if ($ExPropList -eq $null)
		{
			$ExPropList = @()
		}
		if (![String]::IsNullOrEmpty($ItemClass))
		{
			$ItemClassProp = Get-TaggedProperty -DataType "String" -Id "0x001A" -Value $ItemClass
			$ExPropList += $ItemClassProp
		}
		$ExPropList += $SentFlag
		$ExPropList += $SentTime
		$ExPropList += $RcvdTime
		$NewMessage = Get-MessageJSONFormat -Subject $Subject -Body $Body -SenderEmailAddress $SenderEmailAddress -Attachments $Attachments -ToRecipients $ToRecipients -SentDate $SentDate -ExPropList $ExPropList
		if ($FolderPath -ne $null)
		{
			$Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
			
		}
		if ($Folder -ne $null)
		{
			$RequestURL = $Folder.FolderRestURI + "/messages"
			$HttpClient = Get-HTTPClient($MailboxName)
			Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage
		}
		
	}
}

function CreateFlatList
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[psobject]
		$EmailAddress
	)
	Begin
	{
		
		$FlatListEntry = new-object System.IO.MemoryStream
		$EntryOneOffid = "00000000812B1FA4BEA310199D6E00DD010F540200000190" + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(($EmailAddress.Name + "`0"))).Replace("-", "") + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(("SMTP" + "`0"))).Replace("-", "") + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(($EmailAddress.Address + "`0"))).Replace("-", "")
		$FlatListEntryBytes = HexStringToByteArray($EntryOneOffid)
		$FlatListEntry.Write([BitConverter]::GetBytes($FlatListEntryBytes.Length), 0, 4);
		$FlatListEntry.Write($FlatListEntryBytes, 0, $FlatListEntryBytes.Length);
		$InnerLength += $FlatListEntryBytes.Length
		$Modulsval = $FlatListEntryBytes.Length % 4;
		$PadingValue = 0;
		if ($Modulsval -ne 0)
		{
			$PadingValue = 4 - $Modulsval;
			for ($AddPading = 0; $AddPading -lt $PadingValue; $AddPading++)
			{
				[Byte]$NullValue = 00;
				$FlatlistStream.Write($NullValue, 0, 1);
			}
		}
		$FlatListEntry.Position = 0
		$FlatListEntryBytes = $FlatListEntry.ToArray()
		$FlatList = new-object System.IO.MemoryStream
		$FlatList.Write([BitConverter]::GetBytes(1), 0, 4);
		$FlatList.Write([BitConverter]::GetBytes($FlatListEntryBytes.Length), 0, 4);
		$FlatList.Write($FlatListEntryBytes, 0, $FlatListEntryBytes.Length);
		$FlatList.Position = 0
		return, $FlatList.ToArray()
	}
}


function Send-MessageREST
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[PSCustomObject]
		$Folder,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[String]
		$Subject,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$SenderEmailAddress,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[psobject]
		$ReferanceAttachments,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[psobject]
		$ToRecipients,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[psobject]
		$CCRecipients,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[psobject]
		$BCCRecipients,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[psobject]
		$ExPropList,
		
		[Parameter(Position = 14, Mandatory = $false)]
		[psobject]
		$StandardPropList,
		
		[Parameter(Position = 15, Mandatory = $false)]
		[string]
		$ItemClass,
		
		[Parameter(Position = 16, Mandatory = $false)]
		[switch]
		$SaveToSentItems,
		
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
		$ReplyTo
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		if (![String]::IsNullOrEmpty($ItemClass))
		{
			$ItemClassProp = Get-TaggedProperty -DataType "String" -Id "0x001A" -Value $ItemClass
			if ($ExPropList -eq $null)
			{
				$ExPropList = @()
			}
			$ExPropList += $ItemClassProp
		}
		$SaveToSentFolder = "false"
		if ($SaveToSentItems.IsPresent)
		{
			$SaveToSentFolder = "true"
		}
		$NewMessage = Get-MessageJSONFormat -Subject $Subject -Body $Body -SenderEmailAddress $SenderEmailAddress -Attachments $Attachments -ReferanceAttachments $ReferanceAttachments -ToRecipients $ToRecipients -SentDate $SentDate -ExPropList $ExPropList -CcRecipients $CCRecipients -bccRecipients $BCCRecipients -StandardPropList $StandardPropList -SaveToSentItems $SaveToSentFolder -SendMail -ReplyTo $ReplyTo -RequestReadRecipient $RequestReadRecipient.IsPresent -RequestDeliveryRecipient $RequestDeliveryRecipient.IsPresent
		if ($ShowRequest.IsPresent)
		{
			write-host $NewMessage
		}
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName + "/sendmail"
		$HttpClient = Get-HTTPClient($MailboxName)
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage
		
	}
}

function New-HolidayEvent
{
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
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		if ([String]::IsNullOrEmpty($TimeZone))
		{
			$TimeZone = [TimeZoneInfo]::Local.Id
		}
		if ($AccessToken.resource -eq "https://graph.microsoft.com")
		{
			$isAllDay = Get-ItemProp -Name isAllDay -NoQuotes -Value true
		}
		else
		{
			$isAllDay = Get-ItemProp -Name IsAllDay -NoQuotes -Value true
		}
		if ($Calendar -eq $null)
		{
			$Calendar = Get-DefaultCalendarFolder -MailboxName $MailboxName -AccessToken $AccessToken
		}
		if (![String]::IsNullOrEmpty($CalendarName))
		{
			$Calendar = Get-CalendarFolder -MailboxName $MailboxName -AccessToken $AccessToken -FolderName $CalendarName
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
		$NewMessage = Get-EventJSONFormat -Subject $Subject -Body $Body -SenderEmailAddress $SenderEmailAddress -Start $Day.Date -End $Day.Date.AddDays(1) -TimeZone $TimeZone -Attachments $Attachments -ReferanceAttachments $ReferanceAttachments -Attendees $Attendees -SentDate $SentDate -ExPropList $ExPropList -StandardPropList $StandardPropList -ReplyTo $ReplyTo -RequestReadRecipient $RequestReadRecipient.IsPresent -RequestDeliveryRecipient $RequestDeliveryRecipient.IsPresent -Recurrence $Recurrence
		if ($ShowRequest.IsPresent)
		{
			write-host $NewMessage
		}
		
		if ($Group.IsPresent)
		{
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "groups"
			$ModernGroup = Get-ModernGroups -MailboxName $MailboxName -GroupName $GroupName -AccessToken $AccessToken
			$RequestURL = $EndPoint + "('" + $ModernGroup.id + "')/events"
		}
		else
		{
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "/" + $MailboxName + "/calendars('" + $Calendar.id + "')/events"
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage
		
	}
}

function New-CalendarEventREST
{
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
		$Start,
		
		[Parameter(Position = 7, Mandatory = $true)]
		[datetime]
		$End,
		
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
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		if ([String]::IsNullOrEmpty($TimeZone))
		{
			$TimeZone = [TimeZoneInfo]::Local.Id
		}
		if ($Calendar -eq $null)
		{
			$Calendar = Get-DefaultCalendarFolder -MailboxName $MailboxName -AccessToken $AccessToken
		}
		if (![String]::IsNullOrEmpty($CalendarName))
		{
			$Calendar = Get-CalendarFolder -MailboxName $MailboxName -AccessToken $AccessToken -FolderName $CalendarName
			if ([String]::IsNullOrEmpty($Calendar)) { throw "Error Calendar folder not found check the folder name this is case sensitive" }
		}
		$NewMessage = Get-EventJSONFormat -Subject $Subject -Body $Body -SenderEmailAddress $SenderEmailAddress -Start $Start -End $End -TimeZone $TimeZone -Attachments $Attachments -ReferanceAttachments $ReferanceAttachments -Attendees $Attendees -SentDate $SentDate -ExPropList $ExPropList -StandardPropList $StandardPropList -ReplyTo $ReplyTo -RequestReadRecipient $RequestReadRecipient.IsPresent -RequestDeliveryRecipient $RequestDeliveryRecipient.IsPresent -Recurrence $Recurrence
		if ($ShowRequest.IsPresent)
		{
			write-host $NewMessage
		}
		if ($Group.IsPresent)
		{
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "groups"
			$ModernGroup = Get-ModernGroups -MailboxName $MailboxName -GroupName $GroupName -AccessToken $AccessToken
			$RequestURL = $EndPoint + "('" + $ModernGroup.id + "')/events"
		}
		else
		{
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "/" + $MailboxName + "/calendars('" + $Calendar.id + "')/events"
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage
		
	}
}

function Get-MessageJSONFormat
{
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[String]
		$Subject,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
		$SenderEmailAddress,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[psobject]
		$ReferanceAttachments,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[psobject]
		$ToRecipients,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$CcRecipients,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$bccRecipients,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$SentDate,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[psobject]
		$StandardPropList,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[psobject]
		$ExPropList,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[switch]
		$ShowRequest,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[String]
		$SaveToSentItems,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[switch]
		$SendMail,
		
		[Parameter(Position = 14, Mandatory = $false)]
		[psobject]
		$ReplyTo,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[bool]
		$RequestReadRecipient,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[bool]
		$RequestDeliveryRecipient
	)
	Begin
	{
		$NewMessage = "{" + "`r`n"
		if ($SendMail.IsPresent)
		{
			$NewMessage += "  `"Message`" : {" + "`r`n"
		}
		if (![String]::IsNullOrEmpty($Subject))
		{
			$NewMessage += "`"Subject`": `"" + $Subject + "`"" + "`r`n"
		}
		if ($SenderEmailAddress -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Sender`":{" + "`r`n"
			$NewMessage += " `"EmailAddress`":{" + "`r`n"
			$NewMessage += "  `"Name`":`"" + $SenderEmailAddress.Name + "`"," + "`r`n"
			$NewMessage += "  `"Address`":`"" + $SenderEmailAddress.Address + "`"" + "`r`n"
			$NewMessage += "}}" + "`r`n"
		}
		if (![String]::IsNullOrEmpty($Body))
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Body`": {" + "`r`n"
			$NewMessage += "`"ContentType`": `"HTML`"," + "`r`n"
			$NewMessage += "`"Content`": `"" + $Body + "`"" + "`r`n"
			$NewMessage += "}" + "`r`n"
		}
		
		$toRcpcnt = 0;
		if ($ToRecipients -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"ToRecipients`": [ " + "`r`n"
			foreach ($EmailAddress in $ToRecipients)
			{
				if ($toRcpcnt -gt 0)
				{
					$NewMessage += "      ,{ " + "`r`n"
				}
				else
				{
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$toRcpcnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		$ccRcpcnt = 0
		if ($CcRecipients -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"CcRecipients`": [ " + "`r`n"
			foreach ($EmailAddress in $CcRecipients)
			{
				if ($ccRcpcnt -gt 0)
				{
					$NewMessage += "      ,{ " + "`r`n"
				}
				else
				{
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$ccRcpcnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		$bccRcpcnt = 0
		if ($bccRecipients -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"BccRecipients`": [ " + "`r`n"
			foreach ($EmailAddress in $bccRecipients)
			{
				if ($bccRcpcnt -gt 0)
				{
					$NewMessage += "      ,{ " + "`r`n"
				}
				else
				{
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$bccRcpcnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		$ReplyTocnt = 0
		if ($ReplyTo -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"ReplyTo`": [ " + "`r`n"
			foreach ($EmailAddress in $ReplyTo)
			{
				if ($ReplyTocnt -gt 0)
				{
					$NewMessage += "      ,{ " + "`r`n"
				}
				else
				{
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$ReplyTocnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		if ($RequestDeliveryRecipient)
		{
			$NewMessage += ",`"IsDeliveryReceiptRequested`": true`r`n"
		}
		if ($RequestReadRecipient)
		{
			$NewMessage += ",`"IsReadReceiptRequested`": true `r`n"
		}
		if ($StandardPropList -ne $null)
		{
			foreach ($StandardProp in $StandardPropList)
			{
				if ($NewMessage.Length -gt 5) { $NewMessage += "," }
				switch ($StandardProp.PropertyType)
				{
					"Single" {
						if ($StandardProp.QuoteValue)
						{
							$NewMessage += "`"" + $StandardProp.Name + "`": `"" + $StandardProp.Value + "`"" + "`r`n"
						}
						else
						{
							$NewMessage += "`"" + $StandardProp.Name + "`": " + $StandardProp.Value + "`r`n"
						}
						
						
					}
					"Object"  {
						if ($StandardProp.isArray)
						{
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": [ {" + "`r`n"
						}
						else
						{
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
						}
						$acCount = 0
						foreach ($PropKeyValue in $StandardProp.PropertyList)
						{
							if ($acCount -gt 0)
							{
								$NewMessage += ","
							}
							$NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Name + "`"" + "`r`n"
							$acCount++
						}
						if ($StandardProp.isArray)
						{
							$NewMessage += "}]" + "`r`n"
						}
						else
						{
							$NewMessage += "}" + "`r`n"
						}
						
					}
					"ObjectCollection" {
						if ($StandardProp.isArray)
						{
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": [" + "`r`n"
						}
						else
						{
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
						}
						foreach ($EnclosedStandardProp in $StandardProp.PropertyList)
						{
							$NewMessage += "`"" + $EnclosedStandardProp.PropertyName + "`": {" + "`r`n"
							foreach ($PropKeyValue in $EnclosedStandardProp.PropertyList)
							{
								$NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Name + "`"," + "`r`n"
							}
							$NewMessage += "}" + "`r`n"
						}
						if ($StandardProp.isArray)
						{
							$NewMessage += "]" + "`r`n"
						}
						else
						{
							$NewMessage += "}" + "`r`n"
						}
					}
					
				}
			}
		}
		$atcnt = 0
		$processAttachments = $false
		if ($Attachments -ne $null) { $processAttachments = $true }
		if ($ReferanceAttachments -ne $null) { $processAttachments = $true }
		if ($processAttachments)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "  `"Attachments`": [ " + "`r`n"
			if ($Attachments -ne $null)
			{
				foreach ($Attachment in $Attachments)
				{
					$Item = Get-Item $Attachment
					if ($atcnt -gt 0)
					{
						$NewMessage += "   ,{" + "`r`n"
					}
					else
					{
						$NewMessage += "    {" + "`r`n"
					}
					$NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.FileAttachment`"," + "`r`n"
					$NewMessage += "     `"Name`": `"" + $Item.Name + "`"," + "`r`n"
					$NewMessage += "     `"ContentBytes`": `" " + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Attachment)) + "`"" + "`r`n"
					$NewMessage += "    } " + "`r`n"
					$atcnt++
				}
			}
			$atcnt = 0
			if ($ReferanceAttachments -ne $null)
			{
				foreach ($Attachment in $ReferanceAttachments)
				{
					if ($atcnt -gt 0)
					{
						$NewMessage += "   ,{" + "`r`n"
					}
					else
					{
						$NewMessage += "    {" + "`r`n"
					}
					$NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.ReferenceAttachment`"," + "`r`n"
					$NewMessage += "     `"Name`": `"" + $Attachment.Name + "`"," + "`r`n"
					$NewMessage += "     `"SourceUrl`": `"" + $Attachment.SourceUrl + "`"," + "`r`n"
					$NewMessage += "     `"ProviderType`": `"" + $Attachment.ProviderType + "`"," + "`r`n"
					$NewMessage += "     `"Permission`": `"" + $Attachment.Permission + "`"," + "`r`n"
					$NewMessage += "     `"IsFolder`": `"" + $Attachment.IsFolder + "`"" + "`r`n"
					$NewMessage += "    } " + "`r`n"
					$atcnt++
				}
			}
			$NewMessage += "  ]" + "`r`n"
		}
		
		if ($ExPropList -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"SingleValueExtendedProperties`": [" + "`r`n"
			$propCount = 0
			foreach ($Property in $ExPropList)
			{
				if ($propCount -eq 0)
				{
					$NewMessage += "{" + "`r`n"
				}
				else
				{
					$NewMessage += ",{" + "`r`n"
				}
				if ($Property.PropertyType -eq "Tagged")
				{
					$NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Id + "`", " + "`r`n"
				}
				else
				{
					if ($Property.Type -eq "String")
					{
						$NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Guid + " Name " + $Property.Id + "`", " + "`r`n"
					}
					else
					{
						$NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Guid + " Id " + $Property.Id + "`", " + "`r`n"
					}
				}
				$NewMessage += "`"Value`":`"" + $Property.Value + "`"" + "`r`n"
				$NewMessage += " } " + "`r`n"
				$propCount++
			}
			$NewMessage += "]" + "`r`n"
		}
		if (![String]::IsNullOrEmpty($SaveToSentItems))
		{
			$NewMessage += "}   ,`"SaveToSentItems`": `"" + $SaveToSentItems.ToLower() + "`"" + "`r`n"
		}
		$NewMessage += "}"
		if ($ShowRequest.IsPresent)
		{
			Write-Host $NewMessage
		}
		return, $NewMessage
	}
}

function Get-EventJSONFormat
{
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[String]
		$Subject,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[datetime]
		$Start,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[datetime]
		$End,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[psobject]
		$SenderEmailAddress,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$ReferanceAttachments,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$Attendees,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[psobject]
		$SentDate,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[psobject]
		$StandardPropList,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[psobject]
		$ExPropList,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[switch]
		$ShowRequest,
		
		[Parameter(Position = 15, Mandatory = $false)]
		[psobject]
		$ReplyTo,
		
		[Parameter(Position = 16, Mandatory = $false)]
		[bool]
		$RequestReadRecipient,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[bool]
		$RequestDeliveryRecipient,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[String]
		$TimeZone,
		
		[Parameter(Position = 19, Mandatory = $false)]
		[psobject]
		$Recurrence
	)
	Begin
	{
		$NewMessage = "{" + "`r`n"
		if (![String]::IsNullOrEmpty($Subject))
		{
			$NewMessage += "`"Subject`": `"" + $Subject + "`"" + "`r`n"
		}
		if ($Start -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Start`": {   `"DateTime`":`"" + $Start.ToString("yyyy-MM-ddTHH:mm:ss") + "`"," + "`r`n"
			$NewMessage += "  `"TimeZone`":`"" + $TimeZone + "`"}" + "`r`n"
		}
		if ($End -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"End`": {   `"DateTime`":`"" + $End.ToString("yyyy-MM-ddTHH:mm:ss") + "`"," + "`r`n"
			$NewMessage += "  `"TimeZone`":`"" + $TimeZone + "`"}" + "`r`n"
		}
		if ($SenderEmailAddress -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Sender`":{" + "`r`n"
			$NewMessage += " `"EmailAddress`":{" + "`r`n"
			$NewMessage += "  `"Name`":`"" + $SenderEmailAddress.Name + "`"," + "`r`n"
			$NewMessage += "  `"Address`":`"" + $SenderEmailAddress.Address + "`"" + "`r`n"
			$NewMessage += "}}" + "`r`n"
		}
		if (![String]::IsNullOrEmpty($Body))
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Body`": {" + "`r`n"
			$NewMessage += "`"ContentType`": `"HTML`"," + "`r`n"
			$NewMessage += "`"Content`": `"" + $Body + "`"" + "`r`n"
			$NewMessage += "}" + "`r`n"
		}
		
		$toRcpcnt = 0;
		if ($Attendees -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Attendees`": [ " + "`r`n"
			foreach ($Attendee in $Attendees)
			{
				if ($toRcpcnt -gt 0)
				{
					$NewMessage += "      ,{ " + "`r`n"
				}
				else
				{
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $Attendee.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $Attendee.Address + "`"" + "`r`n"
				$NewMessage += "}," + "`r`n"
				$NewMessage += "  `"Type`":`"" + $Attendee.Type + "`"" + " }" + "`r`n"
				$toRcpcnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		if ($Recurrence -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"Recurrence`": { " + "`r`n"
			$NewMessage += "`"Pattern`": { " + "`r`n"
			$NewMessage += "  `"Type`":`"" + $Recurrence.Pattern.Type + "`"," + "`r`n"
			$NewMessage += "  `"Interval`":`"" + $Recurrence.Pattern.Interval + "`"," + "`r`n"
			$NewMessage += "  `"Month`":`"" + $Recurrence.Pattern.Month + "`"," + "`r`n"
			$NewMessage += "  `"DayOfMonth`":`"" + $Recurrence.Pattern.DayOfMonth + "`"," + "`r`n"
			if ($Recurrence.Pattern.DaysOfWeek -ne $null)
			{
				$NewMessage += "  `"DaysOfWeek`":`[" + "`r`n"
				$first = $true;
				foreach ($day in $Recurrence.Pattern.DaysOfWeek)
				{
					if ($first)
					{
						$NewMessage += " `"" + $day + "`"`r`n"
					}
					else
					{
						$NewMessage += ",`"" + $day + "`"`r`n"
					}
					
				}
				$NewMessage += "  ]," + "`r`n"
			}
			$NewMessage += "  `"FirstDayOfWeek`":`"" + $Recurrence.Pattern.FirstDayOfWeek + "`"," + "`r`n"
			$NewMessage += "  `"Index`":`"" + $Recurrence.Pattern.Index + "`"" + "`r`n"
			$NewMessage += "  }," + "`r`n"
			$NewMessage += "`"Range`": { " + "`r`n"
			$NewMessage += "  `"Type`":`"" + $Recurrence.Range.Type + "`"," + "`r`n"
			$NewMessage += "  `"StartDate`":`"" + $Recurrence.Range.StartDate + "`"," + "`r`n"
			$NewMessage += "  `"EndDate`":`"" + $Recurrence.Range.EndDate + "`"," + "`r`n"
			$NewMessage += "  `"RecurrenceTimeZone`":`"" + $Recurrence.RecurrenceTimeZone + "`"," + "`r`n"
			$NewMessage += "  `"NumberOfOccurrences`":`"" + $Recurrence.Range.NumberOfOccurrences + "`"" + "`r`n"
			$NewMessage += "  }" + "`r`n"
			$NewMessage += "  }" + "`r`n"
		}
		$ReplyTocnt = 0
		if ($ReplyTo -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"ReplyTo`": [ " + "`r`n"
			foreach ($EmailAddress in $ReplyTo)
			{
				if ($ReplyTocnt -gt 0)
				{
					$NewMessage += "      ,{ " + "`r`n"
				}
				else
				{
					$NewMessage += "      { " + "`r`n"
				}
				$NewMessage += " `"EmailAddress`":{" + "`r`n"
				$NewMessage += "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
				$NewMessage += "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
				$NewMessage += "}}" + "`r`n"
				$ReplyTocnt++
			}
			$NewMessage += "  ]" + "`r`n"
		}
		if ($RequestDeliveryRecipient)
		{
			$NewMessage += ",`"IsDeliveryReceiptRequested`": true`r`n"
		}
		if ($RequestReadRecipient)
		{
			$NewMessage += ",`"IsReadReceiptRequested`": true `r`n"
		}
		if ($StandardPropList -ne $null)
		{
			foreach ($StandardProp in $StandardPropList)
			{
				if ($NewMessage.Length -gt 5) { $NewMessage += "," }
				switch ($StandardProp.PropertyType)
				{
					"Single" {
						if ($StandardProp.QuoteValue)
						{
							$NewMessage += "`"" + $StandardProp.Name + "`": `"" + $StandardProp.Value + "`"" + "`r`n"
						}
						else
						{
							$NewMessage += "`"" + $StandardProp.Name + "`": " + $StandardProp.Value + "`r`n"
						}
						
						
					}
					"Object"  {
						if ($StandardProp.isArray)
						{
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": [ {" + "`r`n"
						}
						else
						{
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
						}
						$acCount = 0
						foreach ($PropKeyValue in $StandardProp.PropertyList)
						{
							if ($acCount -gt 0)
							{
								$NewMessage += ","
							}
							$NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Name + "`"" + "`r`n"
							$acCount++
						}
						if ($StandardProp.isArray)
						{
							$NewMessage += "}]" + "`r`n"
						}
						else
						{
							$NewMessage += "}" + "`r`n"
						}
						
					}
					"ObjectCollection" {
						if ($StandardProp.isArray)
						{
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": [" + "`r`n"
						}
						else
						{
							$NewMessage += "`"" + $StandardProp.PropertyName + "`": {" + "`r`n"
						}
						foreach ($EnclosedStandardProp in $StandardProp.PropertyList)
						{
							$NewMessage += "`"" + $EnclosedStandardProp.PropertyName + "`": {" + "`r`n"
							foreach ($PropKeyValue in $EnclosedStandardProp.PropertyList)
							{
								$NewMessage += "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Name + "`"," + "`r`n"
							}
							$NewMessage += "}" + "`r`n"
						}
						if ($StandardProp.isArray)
						{
							$NewMessage += "]" + "`r`n"
						}
						else
						{
							$NewMessage += "}" + "`r`n"
						}
					}
					
				}
			}
		}
		$atcnt = 0
		$processAttachments = $false
		if ($Attachments -ne $null) { $processAttachments = $true }
		if ($ReferanceAttachments -ne $null) { $processAttachments = $true }
		if ($processAttachments)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "  `"Attachments`": [ " + "`r`n"
			if ($Attachments -ne $null)
			{
				foreach ($Attachment in $Attachments)
				{
					$Item = Get-Item $Attachment
					if ($atcnt -gt 0)
					{
						$NewMessage += "   ,{" + "`r`n"
					}
					else
					{
						$NewMessage += "    {" + "`r`n"
					}
					$NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.FileAttachment`"," + "`r`n"
					$NewMessage += "     `"Name`": `"" + $Item.Name + "`"," + "`r`n"
					$NewMessage += "     `"ContentBytes`": `" " + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Attachment)) + "`"" + "`r`n"
					$NewMessage += "    } " + "`r`n"
					$atcnt++
				}
			}
			$atcnt = 0
			if ($ReferanceAttachments -ne $null)
			{
				foreach ($Attachment in $ReferanceAttachments)
				{
					if ($atcnt -gt 0)
					{
						$NewMessage += "   ,{" + "`r`n"
					}
					else
					{
						$NewMessage += "    {" + "`r`n"
					}
					$NewMessage += "     `"@odata.type`": `"#Microsoft.OutlookServices.ReferenceAttachment`"," + "`r`n"
					$NewMessage += "     `"Name`": `"" + $Attachment.Name + "`"," + "`r`n"
					$NewMessage += "     `"SourceUrl`": `"" + $Attachment.SourceUrl + "`"," + "`r`n"
					$NewMessage += "     `"ProviderType`": `"" + $Attachment.ProviderType + "`"," + "`r`n"
					$NewMessage += "     `"Permission`": `"" + $Attachment.Permission + "`"," + "`r`n"
					$NewMessage += "     `"IsFolder`": `"" + $Attachment.IsFolder + "`"" + "`r`n"
					$NewMessage += "    } " + "`r`n"
					$atcnt++
				}
			}
			$NewMessage += "  ]" + "`r`n"
		}
		
		if ($ExPropList -ne $null)
		{
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"SingleValueExtendedProperties`": [" + "`r`n"
			$propCount = 0
			foreach ($Property in $ExPropList)
			{
				if ($propCount -eq 0)
				{
					$NewMessage += "{" + "`r`n"
				}
				else
				{
					$NewMessage += ",{" + "`r`n"
				}
				if ($Property.PropertyType -eq "Tagged")
				{
					$NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Id + "`", " + "`r`n"
				}
				else
				{
					if ($Property.Type -eq "String")
					{
						$NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Guid + " Name " + $Property.Id + "`", " + "`r`n"
					}
					else
					{
						$NewMessage += "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Guid + " Id " + $Property.Id + "`", " + "`r`n"
					}
				}
				$NewMessage += "`"Value`":`"" + $Property.Value + "`"" + "`r`n"
				$NewMessage += " } " + "`r`n"
				$propCount++
			}
			$NewMessage += "]" + "`r`n"
		}
		if (![String]::IsNullOrEmpty($SaveToSentItems))
		{
			$NewMessage += "}   ,`"SaveToSentItems`": `"" + $SaveToSentItems.ToLower() + "`"" + "`r`n"
		}
		$NewMessage += "}"
		if ($ShowRequest.IsPresent)
		{
			Write-Host $NewMessage
		}
		return, $NewMessage
	}
}
#endregion