function Invoke-EXRReadEmail
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,

		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$ItemRESTURI
	)
	Process
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$msMessage = Get-EXREmail -MailboxName $MailboxName -ItemRESTURI $ItemRESTURI -AccessToken $AccessToken
        $msgform = new-object System.Windows.Forms.form 
        $msgform.Text = $msMessage.Subject
        $msgform.size = new-object System.Drawing.Size(1000,800) 
        

        # Add Message From Lable
        $miMessageTolableBox = new-object System.Windows.Forms.Label
        $miMessageTolableBox.Location = new-object System.Drawing.Size(20,20) 
        $miMessageTolableBox.size = new-object System.Drawing.Size(80,20) 
        $miMessageTolableBox.Text = "To"
        $msgform.controls.Add($miMessageTolableBox) 

        # Add MessageID Lable
        $miMessageSentlableBox = new-object System.Windows.Forms.Label
        $miMessageSentlableBox.Location = new-object System.Drawing.Size(20,40) 
        $miMessageSentlableBox.size = new-object System.Drawing.Size(80,20) 
        $miMessageSentlableBox.Text = "From"
        $msgform.controls.Add($miMessageSentlableBox) 

        # Add Message Subject Lable
        $miMessageSubjectlableBox = new-object System.Windows.Forms.Label
        $miMessageSubjectlableBox.Location = new-object System.Drawing.Size(20,60) 
        $miMessageSubjectlableBox.size = new-object System.Drawing.Size(80,20) 
        $miMessageSubjectlableBox.Text = "Subject"
        $msgform.controls.Add($miMessageSubjectlableBox) 

        # Add Message To
        $miMessageTotextlabelBox = new-object System.Windows.Forms.Label
        $miMessageTotextlabelBox.Location = new-object System.Drawing.Size(100,20) 
        $miMessageTotextlabelBox.size = new-object System.Drawing.Size(400,20) 
        $msgform.controls.Add($miMessageTotextlabelBox) 
        $ToRecips = "";
        foreach($torcp in $msMessage.toRecipients){
            $ToRecips += $torcp.emailAddress.address.ToString() + ";"
        }
        $miMessageTotextlabelBox.Text = $ToRecips

        # Add Message From
        $miMessageSenttextlabelBox = new-object System.Windows.Forms.Label
        $miMessageSenttextlabelBox.Location = new-object System.Drawing.Size(100,40) 
        $miMessageSenttextlabelBox.size = new-object System.Drawing.Size(600,20) 
        $msgform.controls.Add($miMessageSenttextlabelBox) 
        $miMessageSenttextlabelBox.Text = $msMessage.sender.emailAddress.name.ToString() + " (" + $msMessage.sender.emailAddress.address.ToString() + ")" 

        # Add Message Subject 
        $miMessageSubjecttextlabelBox = new-object System.Windows.Forms.Label
        $miMessageSubjecttextlabelBox.Location = new-object System.Drawing.Size(100,60) 
        $miMessageSubjecttextlabelBox.size = new-object System.Drawing.Size(600,20) 
        $msgform.controls.Add($miMessageSubjecttextlabelBox) 
        $miMessageSubjecttextlabelBox.Text  = $msMessage.Subject.ToString()

        # Add Message body 
        $miMessageBodytextlabelBox = new-object System.Windows.Forms.WebBrowser
        $miMessageBodytextlabelBox.Location = new-object System.Drawing.Size(100,80) 
        $miMessageBodytextlabelBox.size = new-object System.Drawing.Size(900,550) 
        $miMessageBodytextlabelBox.AutoSize = $true
        $miMessageBodytextlabelBox.DocumentText = $msMessage.Body.Content
        $msgform.controls.Add($miMessageBodytextlabelBox) 

        # Add Message Attachments Lable
        $miMessageAttachmentslableBox = new-object System.Windows.Forms.Label
        $miMessageAttachmentslableBox.Location = new-object System.Drawing.Size(20,645) 
        $miMessageAttachmentslableBox.size = new-object System.Drawing.Size(80,20) 
        $miMessageAttachmentslableBox.Text = "Attachments"
        $msgform.controls.Add($miMessageAttachmentslableBox) 

        $miMessageAttachmentslableBox1 = new-object System.Windows.Forms.Label
        $miMessageAttachmentslableBox1.Location = new-object System.Drawing.Size(100,645) 
        $miMessageAttachmentslableBox1.size = new-object System.Drawing.Size(600,20) 
        $miMessageAttachmentslableBox1.Text = ""
        $msgform.Controls.Add($miMessageAttachmentslableBox1) 

        
        $exButton4 = new-object System.Windows.Forms.Button
        $exButton4.Location = new-object System.Drawing.Size(10,665)
        $exButton4.Size = new-object System.Drawing.Size(150,20)
        $exButton4.Text = "Download Attachments"
        $exButton4.Enabled = $false
        $exButton4.Add_Click({DownloadAttachments})
        $msgform.Controls.Add($exButton4)
        
        $attname = ""
        if ($msMessage.hasattachments){
            write-host "Attachment"
            $exButton4.Enabled = $true
            $Attachments = Get-EXRAttachments -MailboxName $MailboxName -AccessToken $AccessToken -ItemURI $ItemRESTURI -MetaData 
            foreach($attach in $Attachments)
            {			
                $attname = $attname + $attach.Name.ToString() + "; "
            }
        }
        $miMessageAttachmentslableBox1.Text = $attname
        # Add Download Button

        $msgform.autoscroll = $true
        $msgform.ShowDialog()
		
	}
}


