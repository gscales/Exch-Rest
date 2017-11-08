function Get-EXRMessageJSONFormat
{
	[CmdletBinding()]
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
				if($Property.Value -eq "null"){
					$NewMessage += "`"Value`":null" + "`r`n"
				}
				else{
					$NewMessage += "`"Value`":`"" + $Property.Value + "`"" + "`r`n"
				}				
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
