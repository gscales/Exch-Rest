function Invoke-EXRNewMessagesForm{
	[CmdletBinding()]
	param (
		
		[Parameter(Position = 0, Mandatory = $false)]
		[String]
		$MailboxName,		
	
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken		
	)
	Process
	{

	$script:newmsgform = new-object System.Windows.Forms.form 
	$script:newmsgform.Text = $MailboxName
	$script:newmsgform.size = new-object System.Drawing.Size(1000,800) 

	# Add Message To Lable
	$miMessageTolableBox = new-object System.Windows.Forms.Label
	$miMessageTolableBox.Location = new-object System.Drawing.Size(20,20) 
	$miMessageTolableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageTolableBox.Text = "To"
	$script:newmsgform.controls.Add($miMessageTolableBox) 

	# Add Message Subject Lable
	$miMessageSubjectlableBox = new-object System.Windows.Forms.Label
	$miMessageSubjectlableBox.Location = new-object System.Drawing.Size(20,65) 
	$miMessageSubjectlableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageSubjectlableBox.Text = "Subject"
	$script:newmsgform.controls.Add($miMessageSubjectlableBox) 

	# Add Message To
	$miMessageTotextlabelBox = new-object System.Windows.Forms.TextBox
	$miMessageTotextlabelBox.Location = new-object System.Drawing.Size(100,20) 
	$miMessageTotextlabelBox.size = new-object System.Drawing.Size(400,20) 
	$script:newmsgform.controls.Add($miMessageTotextlabelBox) 

	# Add Message Subject 
	$miMessageSubjecttextlabelBox = new-object System.Windows.Forms.TextBox
	$miMessageSubjecttextlabelBox.Location = new-object System.Drawing.Size(100,65) 
	$miMessageSubjecttextlabelBox.size = new-object System.Drawing.Size(600,20) 
	$script:newmsgform.controls.Add($miMessageSubjecttextlabelBox) 


	# Add Message body 
	$miMessageBodytextlabelBox = new-object System.Windows.Forms.RichTextBox
	$miMessageBodytextlabelBox.Location = new-object System.Drawing.Size(100,100) 
	$miMessageBodytextlabelBox.size = new-object System.Drawing.Size(600,350) 
	$script:newmsgform.controls.Add($miMessageBodytextlabelBox) 

	# Add Message Attachments Lable
	$miMessageAttachmentslableBox = new-object System.Windows.Forms.Label
	$miMessageAttachmentslableBox.Location = new-object System.Drawing.Size(20,460) 
	$miMessageAttachmentslableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageAttachmentslableBox.Text = "Attachments"
	$script:newmsgform.controls.Add($miMessageAttachmentslableBox) 

	$miMessageAttachmentslableBox1 = new-object System.Windows.Forms.Label
	$miMessageAttachmentslableBox1.Location = new-object System.Drawing.Size(100,460) 
	$miMessageAttachmentslableBox1.size = new-object System.Drawing.Size(600,20) 
	$miMessageAttachmentslableBox1.Text = ""
	$script:newmsgform.Controls.Add($miMessageAttachmentslableBox1) 

	$exButton7 = new-object System.Windows.Forms.Button
	$exButton7.Location = new-object System.Drawing.Size(95,520)
	$exButton7.Size = new-object System.Drawing.Size(125,20)
	$exButton7.Text = "Send Message"
	$exButton7.Add_Click({
		Send-EXRMessageREST -MailboxName $MailboxName -AccessToken $AccessToken -ToRecipients @(New-EXREmailAddress -Address $miMessageTotextlabelBox.Text) -Subject $miMessageSubjecttextlabelBox.Text -Body 		$miMessageBodytextlabelBox.Text -Attachments $script:Attachments
		$script:newmsgform.close()
	})
	$script:newmsgform.Controls.Add($exButton7)

	$exButton4 = new-object System.Windows.Forms.Button
	$exButton4.Location = new-object System.Drawing.Size(95,490)
	$exButton4.Size = new-object System.Drawing.Size(150,20)
	$exButton4.Text = "Add Attachment"
	$exButton4.Enabled = $true
	$exButton4.Add_Click({SelectAttachments})
	
	$script:Attachments = @()

	$script:newmsgform.Controls.Add($exButton4)
	$script:newmsgform.autoscroll = $true
	$script:newmsgform.ShowDialog()

}
}



function SelectAttachments{
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Multiselect = $true 	
}
 
[void]$FileBrowser.ShowDialog()
foreach($File in $FileBrowser.FileNames){
	$script:Attachments += $File
	 $attname += $File + " "
}
$miMessageAttachmentslableBox1.Text = $attname

}