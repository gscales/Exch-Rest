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
