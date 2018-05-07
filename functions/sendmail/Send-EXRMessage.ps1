function Send-EXRMessage
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		
		[Parameter(Position = 4, Mandatory = $true)]
		[String]
		$Subject,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$From,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$Attachment,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[psobject]
		$To,

		[Parameter(Position = 11, Mandatory = $false)]
		[psobject]
		$CC,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[psobject]
		$BCC,
		
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
		$SaveToSentFolder = "false"
		if ($SaveToSentItems.IsPresent)
		{
			$SaveToSentFolder = "true"
		}
		$Attachments = @()
		if(![String]::IsNullOrEmpty($Attachment)){
			$Attachments += (Resolve-Path $Attachment).Path
		}
		$ToRecipients = @()
		if(![String]::IsNullOrEmpty($To)){
			$ToRecipients += (New-EXREmailAddress -Address $To)
		}
		$CCRecipients = @()
		if(![String]::IsNullOrEmpty($CC)){
		   $CCRecipients += (New-EXREmailAddress -Address $CC)
		}
		$BCCRecipients = @()
		if(![String]::IsNullOrEmpty($BCC)){
		   $BCCRecipients += (New-EXREmailAddress -Address $BCC)
		}
		$SenderEmailAddress = ""
		if(![String]::IsNullOrEmpty($From)){
			$SenderEmailAddress = (New-EXREmailAddress -Address $From)
		}
		$NewMessage = Get-MessageJSONFormat -Subject $Subject -Body $Body.Replace("`"","\`"") -SenderEmailAddress $SenderEmailAddress -Attachments $Attachments -ReferanceAttachments $ReferanceAttachments -ToRecipients $ToRecipients -SentDate $SentDate -ExPropList $ExPropList -CcRecipients $CCRecipients -bccRecipients $BCCRecipients -StandardPropList $StandardPropList -SaveToSentItems $SaveToSentFolder -SendMail -ReplyTo $ReplyTo -RequestReadRecipient $RequestReadRecipient.IsPresent -RequestDeliveryRecipient $RequestDeliveryRecipient.IsPresent
		if ($ShowRequest.IsPresent)
		{
			write-host $NewMessage
		}
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName + "/sendmail"
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage
		
	}
}
