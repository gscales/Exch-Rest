function Send-ClientMessage
{
	[CmdletBinding()]
	Param (
		
	)
	Send-EXRMessageREST -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ToRecipients @(New-EXREmailAddress -Address $miMessageTotextlabelBox.Text) -Subject $miMessageSubjecttextlabelBox.Text -Body $miMessageBodytextlabelBox.Text -Attachments $script:Attachments
	
	$script:newmsgform.close()
	
}