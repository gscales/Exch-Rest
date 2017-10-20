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
