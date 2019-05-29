function Send-EXRVoiceMail {
    [CmdletBinding()]
    param( 
        [Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 3, Mandatory = $true)] [string]$Mp3FileName,
        [Parameter(Position = 4, Mandatory = $true)] [string]$ToAddress, 
        [Parameter(Position = 5, Mandatory = $false)] [string]$Transcription 


    )
    Process {
        $shell = New-Object -COMObject Shell.Application
        $folder = Split-Path $Mp3FileName
        $file = Split-Path $Mp3FileName -Leaf
        $shellfolder = $shell.Namespace($folder)
        $shellfile = $shellfolder.ParseName($file)
        $dt = [DateTime]::ParseExact($shellfolder.GetDetailsOf($shellfile, 27), "HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture);
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        } 
        $UserResult = Get-EXRUsers -filter ("mail eq '" + $MailboxName + "'")
        $VoiceMailSuject = "Voice Mail (" + $dt.TimeOfDay.TotalSeconds + " seconds)"
        $duration = $dt.TimeOfDay.TotalSeconds
        $voiceMailFrom = $UserResult.displayName
        if($UserResult.businessPhones.count -gt 0){
            $callerId = $UserResult.businessPhones[0]
        }        
        $jobTitle = $UserResult.jobTitle.ToString()
        $Company = $UserResult.companyName
        $BusinessPhone = $callerId
        $emailAddress = $UserResult.mail
        $MobilePhone = $UserResult.mobilePhone.ToString()       
        $BodyHtml = "<html><head><META HTTP-EQUIV=`"Content-Type`" CONTENT=`"text/html; charset=us-ascii`">"
        $BodyHtml += "<style type=`"text/css`"> a:link { color: #3399ff; } a:visited { color: #3366cc; } a:active { color: #ff9900; } </style>"
        $BodyHtml += "</head><body><style type=`"text/css`"> a:link { color: #3399ff; } a:visited { color: #3366cc; } a:active { color: #ff9900; } </style>"
        $BodyHtml += "<div style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`"><div id=`"UM-call-info`" lang=`"en`">"
        $BodyHtml += "<div style=`"font-family: Arial; font-size: 10pt; color:#000066; font-weight: bold;`">You received a voice mail from " + $voiceMailFrom + " at " + $MailboxName + "</div>"
        $BodyHtml += "<br><table border=`"0`" width=`"100%`">"
        $BodyHtml += "<tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
        $BodyHtml += "Company:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
        $BodyHtml += $Company + "</td></tr>"
        $BodyHtml +=  "<tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
        $BodyHtml += "Title:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
        $BodyHtml += $jobTitle + "</td></tr><tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
        $BodyHtml += "Work:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
        $BodyHtml += "<a style=`"color: #3399ff; `" dir=`"ltr`" href=`"tel:" + $BusinessPhone + "`">" + $BusinessPhone + "</a></td></tr>"
        $BodyHtml += "<tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
        $BodyHtml += "Mobile:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
        $BodyHtml += "<a style=`"color: #3399ff; `" dir=`"ltr`" href=`"tel:&#43;" + $MobilePhone + "`">&#43;" + $MobilePhone + "</a></td></tr>"
        $BodyHtml += "</table></div></div></body></html>"
        $ToRecp = "" | Select-Object Name, Address
        $ToRecp.Name = $ToAddress 
        $ToRecp.Address = $ToAddress
        $SenderAddress = "" | Select-Object Name, Address
        $SenderAddress.Name = $MailboxName 
        $SenderAddress.Address = $MailboxName
        $ItemClassProp = "" | Select Id,DataType,PropertyType,Value
        $ItemClassProp.id = "0x001A"
        $ItemClassProp.DataType = "String"
        $ItemClassProp.PropertyType = "Tagged"
        $ItemClassProp.Value = "IPM.Note.Microsoft.Voicemail.UM.CA"
        $VoiceMailLength = "" | Select Id,DataType,PropertyType,Type,Guid,Value
        $VoiceMailLength.id = "0x6801"
        $VoiceMailLength.DataType = "Integer"
        $VoiceMailLength.Guid = "{00020328-0000-0000-c000-000000000046}"
        $VoiceMailLength.PropertyType = "Named"
        $VoiceMailLength.Type = "Id"
        $VoiceMailLength.Value = $dt.TimeOfDay.TotalSeconds
        $VoiceMessageConfidenceLevel = "" | Select Id,DataType,Type,PropertyType,Guid,Value
        $VoiceMessageConfidenceLevel.Id = "X-VoiceMessageConfidenceLevel"
        $VoiceMessageConfidenceLevel.DataType = "String"
        $VoiceMessageConfidenceLevel.Guid = "{00020386-0000-0000-C000-000000000046}"
        $VoiceMessageConfidenceLevel.PropertyType = "Named"
        $VoiceMessageConfidenceLevel.Value = "high"
        $VoiceMessageConfidenceLevel.Type = "String"
        $VoiceMessageTranscription = "" | Select Id,DataType,Type,PropertyType,Guid,Value
        $VoiceMessageTranscription.Id = "X-VoiceMessageTranscription"
        $VoiceMessageTranscription.DataType = "String"
        $VoiceMessageTranscription.Guid = "{00020386-0000-0000-C000-000000000046}"
        $VoiceMessageTranscription.PropertyType = "Named"
        $VoiceMessageTranscription.Value = $Transcription
        $VoiceMessageTranscription.Type = "String"
        $exProp = @()
        $exProp += $ItemClassProp
        $exProp += $VoiceMailLength
        $exProp += $VoiceMessageConfidenceLevel
        $exProp += $VoiceMessageTranscription
        $Attachment = "" | Select name,contentBytes
        $Attachment.name = "audio.mp3"
        $Attachment.contentBytes = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Mp3FileName))
        $NewMessage = Get-MessageJSONFormat -Subject $VoiceMailSuject -Body $BodyHtml.Replace("`"","\`"") -SenderEmailAddress $SenderAddress -Attachments @($Attachment) -ToRecipients @($ToRecp) -SaveToSentItems "true" -SendMail -ExPropList $exProp
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName + "/sendmail"
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage
    }
}