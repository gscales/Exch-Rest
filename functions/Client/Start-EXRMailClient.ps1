function Start-EXRMailClient
{
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
		[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
		[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
		$Script:form = new-object System.Windows.Forms.form 
		$Script:Treeinfo = @{ }
		$mbtable = New-Object System.Data.DataTable
		$mbtable.TableName = "Folder Item"
		$mbtable.Columns.Add("From")
		$mbtable.Columns.Add("Subject")
		$mbtable.Columns.Add("Recieved",[DATETIME])
		$mbtable.Columns.Add("Size",[INT64])
		$mbtable.Columns.Add("ID")
		$mbtable.Columns.Add("hasAttachments")

		# Add Email Address 
		$emEmailAddressTextBox = new-object System.Windows.Forms.TextBox 
		$emEmailAddressTextBox.Location = new-object System.Drawing.Size(130,20) 
		$emEmailAddressTextBox.size = new-object System.Drawing.Size(300,20) 
		$emEmailAddressTextBox.Enabled = $true
		$emEmailAddressTextBox.text =""
		$Script:form.controls.Add($emEmailAddressTextBox) 

		# Add  Email Address  Lable
		$emEmailAddresslableBox = new-object System.Windows.Forms.Label
		$emEmailAddresslableBox.Location = new-object System.Drawing.Size(10,20) 
		$emEmailAddresslableBox.size = new-object System.Drawing.Size(120,20) 
		$emEmailAddresslableBox.Text = "Email Address"
		$Script:form.controls.Add($emEmailAddresslableBox) 


		# Add ClientId Box
		$unCASUrlTextBox = new-object System.Windows.Forms.TextBox 
		$unCASUrlTextBox.Location = new-object System.Drawing.Size(130,45) 
		$unCASUrlTextBox.size = new-object System.Drawing.Size(400,20) 
		$unCASUrlTextBox.text = "5471030d-f311-4c5d-91ef-74ca885463a7"
		$unCASUrlTextBox.Enabled = $true
		$Script:form.Controls.Add($unCASUrlTextBox) 

		# Add CASUrl Lable
		$unCASUrllableBox = new-object System.Windows.Forms.Label
		$unCASUrllableBox.Location = new-object System.Drawing.Size(10,45) 
		$unCASUrllableBox.size = new-object System.Drawing.Size(50,20) 
		$unCASUrllableBox.Text = "ClientId"
		$Script:form.Controls.Add($unCASUrllableBox) 

		# Add redirect Box
		$RedirectTextBox = new-object System.Windows.Forms.TextBox 
		$RedirectTextBox.Location = new-object System.Drawing.Size(130,70) 
		$RedirectTextBox.size = new-object System.Drawing.Size(400,20) 
		$RedirectTextBox.text = "urn:ietf:wg:oauth:2.0:oob"
		$RedirectTextBox.Enabled = $true
		$Script:form.Controls.Add($RedirectTextBox) 

		# Add redirect Lable
		$RedirectlableBox = new-object System.Windows.Forms.Label
		$RedirectlableBox.Location = new-object System.Drawing.Size(10,70) 
		$RedirectlableBox.size = new-object System.Drawing.Size(100,20) 
		$RedirectlableBox.Text = "RedirectURL"
		$Script:form.Controls.Add($RedirectlableBox) 


		$exButton1 = new-object System.Windows.Forms.Button
		$exButton1.Location = new-object System.Drawing.Size(10,130)
		$exButton1.Size = new-object System.Drawing.Size(125,20)
		$exButton1.Text = "Open Mailbox"
		$exButton1.Add_Click({OpenMailbox})
		$Script:form.Controls.Add($exButton1)

		# Add Numeric Results

		$neResultCheckNum =  new-object System.Windows.Forms.numericUpDown
		$neResultCheckNum.Location = new-object System.Drawing.Size(250,130)
		$neResultCheckNum.Size = new-object System.Drawing.Size(70,30)
		$neResultCheckNum.Enabled = $true
		$neResultCheckNum.Value = 100
		$neResultCheckNum.Maximum = 10000000000
		$Script:form.Controls.Add($neResultCheckNum)

		$exButton2 = new-object System.Windows.Forms.Button
		$exButton2.Location = new-object System.Drawing.Size(330,130)
		$exButton2.Size = new-object System.Drawing.Size(125,25)
		$exButton2.Text = "Show Message"
		$exButton2.Add_Click({ShowMessage})
		$Script:form.Controls.Add($exButton2)

		$exButton5 = new-object System.Windows.Forms.Button
		$exButton5.Location = new-object System.Drawing.Size(455,130)
		$exButton5.Size = new-object System.Drawing.Size(125,25)
		$exButton5.Text = "Show Header"
		$exButton5.Add_Click({ShowHeader})
		$Script:form.Controls.Add($exButton5)

		$exButton6 = new-object System.Windows.Forms.Button
		$exButton6.Location = new-object System.Drawing.Size(330,155)
		$exButton6.Size = new-object System.Drawing.Size(125,25)
		$exButton6.Text = "New Message"
		$exButton6.Add_Click({NewMessage})
		$Script:form.Controls.Add($exButton6)

		$exButton7 = new-object System.Windows.Forms.Button
		$exButton7.Location = new-object System.Drawing.Size(960,165)
		$exButton7.Size = new-object System.Drawing.Size(90,25)
		$exButton7.Text = "Update"
		$exButton7.Add_Click({GetFolderItems})
		$Script:form.Controls.Add($exButton7)


		# Add Search Lable

		$saSeachBoxLable = new-object System.Windows.Forms.Label
		$saSeachBoxLable.Location = new-object System.Drawing.Size(600,135) 
		$saSeachBoxLable.Size = new-object System.Drawing.Size(170,20) 
		$saSeachBoxLable.Text = "Search by Property"
		$Script:form.controls.Add($saSeachBoxLable) 

		$saNumItemsBoxLable = new-object System.Windows.Forms.Label
		$saNumItemsBoxLable.Location = new-object System.Drawing.Size(160,135) 
		$saNumItemsBoxLable.Size = new-object System.Drawing.Size(170,20) 
		$saNumItemsBoxLable.Text = "Number of Items"
		$Script:form.controls.Add($saNumItemsBoxLable) 

		$seSearchCheck =  new-object System.Windows.Forms.CheckBox
		$seSearchCheck.Location = new-object System.Drawing.Size(585,130)
		$seSearchCheck.Size = new-object System.Drawing.Size(30,25)
		$seSearchCheck.Add_Click({if ($seSearchCheck.Checked -eq $false){
			$sbSearchTextBox.Enabled = $false
			$snSearchPropDrop.Enabled = $false
			}
			else{
				$sbSearchTextBox.Enabled = $true
				$snSearchPropDrop.Enabled = $true
			}
		})
		$Script:form.controls.Add($seSearchCheck)

		#Add Search box
		$snSearchPropDrop = new-object System.Windows.Forms.ComboBox
		$snSearchPropDrop.Location = new-object System.Drawing.Size(585,165)
		$snSearchPropDrop.Size = new-object System.Drawing.Size(150,30)
		$snSearchPropDrop.Items.Add("Subject")
		$snSearchPropDrop.Items.Add("Body")
		$snSearchPropDrop.Items.Add("From")
		$snSearchPropDrop.Enabled = $false
		$Script:form.Controls.Add($snSearchPropDrop)

		# Add Search TextBox
		$sbSearchTextBox = new-object System.Windows.Forms.TextBox 
		$sbSearchTextBox.Location = new-object System.Drawing.Size(750,165) 
		$sbSearchTextBox.size = new-object System.Drawing.Size(200,20) 
		$sbSearchTextBox.Enabled = $false
		$Script:form.controls.Add($sbSearchTextBox) 

		$tvTreView = new-object System.Windows.Forms.TreeView
		$tvTreView.Location = new-object System.Drawing.Size(10,155)  
		$tvTreView.size = new-object System.Drawing.Size(216,400) 
		$tvTreView.Anchor = "Top,left,Bottom"
		$tvTreView.add_AfterSelect({
			$Script:lfFolderID = $this.SelectedNode.tag
			GetFolderItems
			
		})
		$Script:form.Controls.Add($tvTreView)

		# Add DataGrid View

		$dgDataGrid = new-object System.windows.forms.DataGridView
		$dgDataGrid.Location = new-object System.Drawing.Size(250,200) 
		$dgDataGrid.size = new-object System.Drawing.Size(800,600)
		$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
		$dgDataGrid.AllowUserToDeleteRows = $false
		$dgDataGrid.AllowUserToAddRows = $false
		$Script:form.Controls.Add($dgDataGrid)

		$Script:form.Text = "Simple Exchange Mailbox Client"
		$Script:form.size = new-object System.Drawing.Size(1200,800) 
		$Script:form.autoscroll = $true
		$Script:form.Add_Shown({$Script:form.Activate()})
		if ($AccessToken -ne $null){
			 $Script:AccessToken = $AccessToken
			 $emEmailAddressTextBox.Text = $MailboxName
		     OpenMailbox -AccessToken $AccessToken
		}
		$Script:form.ShowDialog()

	}
}
function OpenMailbox(){
		[CmdletBinding()]
		param (
	


		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken

		
	    )
		Process
	{
		$tvTreView.Nodes.Clear()
		$Script:Treeinfo.Clear()
		if($AccessToken -eq $null){
			$Script:AccessToken = Get-EXRAccessToken -MailboxName $emEmailAddressTextBox.Text -ClientId $unCASUrlTextBox.Text -redirectUrl $RedirectTextBox.Text -ResourceURL graph.Microsoft.com
		}
		else{
			$Script:AccessToken = $AccessToken
		}		
	    $rootFolder = Get-EXRRootMailFolder -AccessToken $Script:AccessToken -MailboxName $emEmailAddressTextBox.Text
		if ($ShowFolderSize)
		{
			$PropList = @()
			$FolderSizeProp = Get-EXRTaggedProperty -Id "0x0E08" -DataType Long
			$PropList += $FolderSizeProp
			$Folders = Get-EXRAllMailFolders -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken  -PropList $PropList
		}
		else
		{
			$Folders = Get-EXRAllMailFolders -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken 
		}	
		$Script:Treeinfo = @{ }
		$TNRoot = new-object System.Windows.Forms.TreeNode("Root")
		$TNRoot.Name = "Mailbox"
		$TNRoot.Text = "Mailbox - " + $emEmailAddressTextBox.Text
		$exProgress = 0
		foreach ($ffFolder in $Folders)
		{
			#Process folder here
			$ParentFolderId = $ffFolder.parentFolderId
			$folderName = $ffFolder.displayName
			if ($ShowFolderSize)
			{
				$folderName = $ffFolder.displayName + " (" + [math]::round($ffFolder.singleValueExtendedProperties[0].value /1Mb, 0) + " mb)"
			}
			$TNChild = new-object System.Windows.Forms.TreeNode($ffFolder.Name)
			$TNChild.Name = $folderName
			$TNChild.Text = $folderName
			$TNChild.tag = $ffFolder			
			
			if ($ParentFolderId -eq $rootFolder.Id)
			{
				[void]$TNRoot.Nodes.Add($TNChild)
				$Script:Treeinfo.Add($ffFolder.Id.ToString(), $TNChild)
			}
			else
			{
				$pfFolder = $Script:Treeinfo[$ParentFolderId]
				[void]$pfFolder.Nodes.Add($TNChild)
				if ($Script:Treeinfo.ContainsKey($ffFolder.Id) -eq $false)
				{
					$Script:Treeinfo.Add($ffFolder.Id, $TNChild)
				}
			}
		}
		$Script:clickedFolder = $null
		[void]$tvTreView.Nodes.Add($TNRoot)
		Write-Progress -Activity "Executing Request" -Completed
	}
}

function GetFolderItems(){
	$mbtable.Clear()
	$folder = $Script:lfFolderID 
	if($seSearchCheck.Checked){
			switch($snSearchPropDrop.SelectedItem.ToString()){
				"Subject" {$sfilter = "Subject eq '" + $sbSearchTextBox.Text.ToString() + "'"
				$Items = Get-EXRFolderItems -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ReturnSize -Folder $folder -TopOnly:$true -Top 100 -Filter $sfilter -TrackStatus
			}
				"Body" {$sfilter = "`"Body:'" + $sbSearchTextBox.Text.ToString() + "'`""
			    $Items = Get-EXRFolderItems -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ReturnSize -Folder $folder -TopOnly:$true -Top 100 -Search $sfilter -TrackStatus
			}
				"From" {$sfilter = "`"From:'" + $sbSearchTextBox.Text.ToString() + "'`""
				$Items = Get-EXRFolderItems -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ReturnSize -Folder $folder -TopOnly:$true -Top 100 -Search $sfilter -TrackStatus

			}
			}
			
	}
	else{
		$Items = Get-EXRFolderItems -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ReturnSize -Folder $folder -TopOnly:$true -Top 100 -TrackStatus
	} 
	foreach($mail in $Items){
		if ($mail.sender.emailAddress.name -ne $null){$fnFromName = $mail.sender.emailAddress.name}
		else{$fnFromName = "N/A"}
		if ($mail.Subject -ne $null){$sbSubject = $mail.Subject.ToString()}
		else{$sbSubject = "N/A"}
		if ([bool]($mail.PSobject.Properties.name -match "Size")){
			$mbtable.rows.add($fnFromName,$sbSubject,$mail.receivedDateTime,$mail.Size.ToString(),$mail.ItemRESTURI,$mail.hasAttachments)
		}
		else{
			$mbtable.rows.add($fnFromName,$sbSubject,$mail.receivedDateTime,0,$mail.ItemRESTURI,$mail.hasAttachments)
		}
	}
	$dgDataGrid.DataSource = $mbtable
}

function newMessage($reply){
	$script:newmsgform = new-object System.Windows.Forms.form 
	$script:newmsgform.Text = "New Message"
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
	$exButton7.Add_Click({SendMessage})
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
	$script:newmsgform.Add_Shown({$Script:form.Activate()})
	$script:newmsgform.ShowDialog()

}

function SendMessage(){
	Send-EXRMessageREST -MailboxName $emEmailAddressTextBox.Text  -AccessToken $Script:AccessToken -ToRecipients @(New-EXREmailAddress -Address $miMessageTotextlabelBox.Text) -Subject $miMessageSubjecttextlabelBox.Text -Body $miMessageBodytextlabelBox.Text -Attachments $script:Attachments
	
	$script:newmsgform.close()

}


function showMessage($MessageID){
    $MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$script:msMessage = Get-EXREmail -MailboxName $emEmailAddressTextBox.Text -ItemRESTURI $MessageID -AccessToken $Script:AccessToken
	write-host $MessageID
	$msgform = new-object System.Windows.Forms.form 
	$msgform.Text = $script:msMessage.Subject
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
	if ($script:msMessage.hasattachments){
		write-host "Attachment"
		$exButton4.Enabled = $true
		$Attachments = Get-EXRAttachments -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ItemURI $MessageID -MetaData 
		foreach($attach in $Attachments)
		{			
			$attname = $attname + $attach.Name.ToString() + "; "
		}
	}
	$miMessageAttachmentslableBox1.Text = $attname
	# Add Download Button

	$msgform.autoscroll = $true
	$msgform.Add_Shown({$Script:form.Activate()})
	$msgform.ShowDialog()

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

function downloadattachments{
	$dlfolder = new-object -com shell.application 
	$dlfolderpath = $dlfolder.BrowseForFolder(0,"Download attachments to",0) 
	Get-EXRAttachments -MailboxName $emEmailAddressTextBox.Text -ItemURI $Script:msMessage.ItemRESTURI -MetaData -AccessToken $Script:AccessToken | ForEach-Object{
            $attach = Invoke-EXRDownloadAttachment -MailboxName $emEmailAddressTextBox.Text -AttachmentURI $_.AttachmentRESTURI -AccessToken $Script:AccessToken
           	$fiFile = new-object System.IO.FileStream(($dlfolderpath.Self.Path  + "\" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
            $attachBytes = [System.Convert]::FromBase64String($attach.ContentBytes)   
		    $fiFile.Write($attachBytes, 0, $attachBytes.Length)
		    $fiFile.Close()
		    write-host ("Downloaded Attachment : " + (($dlfolderpath.Self.Path + "\" + $attach.Name.ToString())))
    }
}

function ShowHeader{
    $MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$script:msMessage = Get-EXREmail -MailboxName $emEmailAddressTextBox.Text -ItemRESTURI $MessageID -AccessToken $Script:AccessToken -PropList (Get-EXRTransportHeader)
	write-host $MessageID
	$hdrform = new-object System.Windows.Forms.form 
	$hdrform.Text = $script:msMessage.Subject
	$hdrform.size = new-object System.Drawing.Size(800,600) 
	# Add Message header
	$miMessageHeadertextlabelBox = new-object System.Windows.Forms.RichTextBox
	$miMessageHeadertextlabelBox.Location = new-object System.Drawing.Size(10,10) 
	$miMessageHeadertextlabelBox.size = new-object System.Drawing.Size(800,600) 
	$miMessageHeadertextlabelBox.text = $script:msMessage.PR_TRANSPORT_MESSAGE_HEADERS
	$hdrform.controls.Add($miMessageHeadertextlabelBox) 
	$hdrform.autoscroll = $true
	$hdrform.Add_Shown({$Script:form.Activate()})
	$hdrform.ShowDialog()


}


