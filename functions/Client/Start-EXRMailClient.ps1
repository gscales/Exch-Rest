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
		$mbtable.Columns.Add("Recieved", [DATETIME])
		$mbtable.Columns.Add("Size", [INT64])
		$mbtable.Columns.Add("ID")
		$mbtable.Columns.Add("hasAttachments")
		
		# Add Email Address 
		$emEmailAddressTextBox = new-object System.Windows.Forms.TextBox
		$emEmailAddressTextBox.Location = new-object System.Drawing.Size(130, 20)
		$emEmailAddressTextBox.size = new-object System.Drawing.Size(300, 20)
		$emEmailAddressTextBox.Enabled = $true
		$emEmailAddressTextBox.text = ""
		$Script:form.controls.Add($emEmailAddressTextBox)
		
		# Add  Email Address  Lable
		$emEmailAddresslableBox = new-object System.Windows.Forms.Label
		$emEmailAddresslableBox.Location = new-object System.Drawing.Size(10, 20)
		$emEmailAddresslableBox.size = new-object System.Drawing.Size(120, 20)
		$emEmailAddresslableBox.Text = "Email Address"
		$Script:form.controls.Add($emEmailAddresslableBox)
		
		
		# Add ClientId Box
		$unCASUrlTextBox = new-object System.Windows.Forms.TextBox
		$unCASUrlTextBox.Location = new-object System.Drawing.Size(130, 45)
		$unCASUrlTextBox.size = new-object System.Drawing.Size(400, 20)
		$unCASUrlTextBox.text = "5471030d-f311-4c5d-91ef-74ca885463a7"
		$unCASUrlTextBox.Enabled = $true
		$Script:form.Controls.Add($unCASUrlTextBox)
		
		# Add CASUrl Lable
		$unCASUrllableBox = new-object System.Windows.Forms.Label
		$unCASUrllableBox.Location = new-object System.Drawing.Size(10, 45)
		$unCASUrllableBox.size = new-object System.Drawing.Size(50, 20)
		$unCASUrllableBox.Text = "ClientId"
		$Script:form.Controls.Add($unCASUrllableBox)
		
		# Add redirect Box
		$RedirectTextBox = new-object System.Windows.Forms.TextBox
		$RedirectTextBox.Location = new-object System.Drawing.Size(130, 70)
		$RedirectTextBox.size = new-object System.Drawing.Size(400, 20)
		$RedirectTextBox.text = "urn:ietf:wg:oauth:2.0:oob"
		$RedirectTextBox.Enabled = $true
		$Script:form.Controls.Add($RedirectTextBox)
		
		# Add redirect Lable
		$RedirectlableBox = new-object System.Windows.Forms.Label
		$RedirectlableBox.Location = new-object System.Drawing.Size(10, 70)
		$RedirectlableBox.size = new-object System.Drawing.Size(100, 20)
		$RedirectlableBox.Text = "RedirectURL"
		$Script:form.Controls.Add($RedirectlableBox)
		
		
		$exButton1 = new-object System.Windows.Forms.Button
		$exButton1.Location = new-object System.Drawing.Size(10, 130)
		$exButton1.Size = new-object System.Drawing.Size(125, 20)
		$exButton1.Text = "Open Mailbox"
		$exButton1.Add_Click({ Open-ClientMailbox })
		$Script:form.Controls.Add($exButton1)
		
		# Add Numeric Results
		$neResultCheckNum = new-object System.Windows.Forms.numericUpDown
		$neResultCheckNum.Location = new-object System.Drawing.Size(250, 130)
		$neResultCheckNum.Size = new-object System.Drawing.Size(70, 30)
		$neResultCheckNum.Enabled = $true
		$neResultCheckNum.Value = 100
		$neResultCheckNum.Maximum = 10000000000
		$Script:form.Controls.Add($neResultCheckNum)
		
		$exButton2 = new-object System.Windows.Forms.Button
		$exButton2.Location = new-object System.Drawing.Size(330, 130)
		$exButton2.Size = new-object System.Drawing.Size(125, 25)
		$exButton2.Text = "Show Message"
		$exButton2.Add_Click({ Show-ClientMessage })
		$Script:form.Controls.Add($exButton2)
		
		$exButton5 = new-object System.Windows.Forms.Button
		$exButton5.Location = new-object System.Drawing.Size(455, 130)
		$exButton5.Size = new-object System.Drawing.Size(125, 25)
		$exButton5.Text = "Show Header"
		$exButton5.Add_Click({ Show-ClientHeader })
		$Script:form.Controls.Add($exButton5)
		
		$exButton6 = new-object System.Windows.Forms.Button
		$exButton6.Location = new-object System.Drawing.Size(330, 155)
		$exButton6.Size = new-object System.Drawing.Size(125, 25)
		$exButton6.Text = "New Message"
		$exButton6.Add_Click({ New-ClientMessage })
		$Script:form.Controls.Add($exButton6)
		
		$exButton7 = new-object System.Windows.Forms.Button
		$exButton7.Location = new-object System.Drawing.Size(960, 165)
		$exButton7.Size = new-object System.Drawing.Size(90, 25)
		$exButton7.Text = "Update"
		$exButton7.Add_Click({ Get-ClientFolderItem })
		$Script:form.Controls.Add($exButton7)
		
		
		# Add Search Lable
		$saSeachBoxLable = new-object System.Windows.Forms.Label
		$saSeachBoxLable.Location = new-object System.Drawing.Size(600, 135)
		$saSeachBoxLable.Size = new-object System.Drawing.Size(170, 20)
		$saSeachBoxLable.Text = "Search by Property"
		$Script:form.controls.Add($saSeachBoxLable)
		
		$saNumItemsBoxLable = new-object System.Windows.Forms.Label
		$saNumItemsBoxLable.Location = new-object System.Drawing.Size(160, 135)
		$saNumItemsBoxLable.Size = new-object System.Drawing.Size(170, 20)
		$saNumItemsBoxLable.Text = "Number of Items"
		$Script:form.controls.Add($saNumItemsBoxLable)
		
		$seSearchCheck = new-object System.Windows.Forms.CheckBox
		$seSearchCheck.Location = new-object System.Drawing.Size(585, 130)
		$seSearchCheck.Size = new-object System.Drawing.Size(30, 25)
		$seSearchCheck.Add_Click({
				if ($seSearchCheck.Checked -eq $false)
				{
					$sbSearchTextBox.Enabled = $false
					$snSearchPropDrop.Enabled = $false
				}
				else
				{
					$sbSearchTextBox.Enabled = $true
					$snSearchPropDrop.Enabled = $true
				}
			})
		$Script:form.controls.Add($seSearchCheck)
		
		#Add Search box
		$snSearchPropDrop = new-object System.Windows.Forms.ComboBox
		$snSearchPropDrop.Location = new-object System.Drawing.Size(585, 165)
		$snSearchPropDrop.Size = new-object System.Drawing.Size(150, 30)
		$snSearchPropDrop.Items.Add("Subject")
		$snSearchPropDrop.Items.Add("Body")
		$snSearchPropDrop.Items.Add("From")
		$snSearchPropDrop.Enabled = $false
		$Script:form.Controls.Add($snSearchPropDrop)
		
		# Add Search TextBox
		$sbSearchTextBox = new-object System.Windows.Forms.TextBox
		$sbSearchTextBox.Location = new-object System.Drawing.Size(750, 165)
		$sbSearchTextBox.size = new-object System.Drawing.Size(200, 20)
		$sbSearchTextBox.Enabled = $false
		$Script:form.controls.Add($sbSearchTextBox)
		
		$tvTreView = new-object System.Windows.Forms.TreeView
		$tvTreView.Location = new-object System.Drawing.Size(10, 155)
		$tvTreView.size = new-object System.Drawing.Size(216, 400)
		$tvTreView.Anchor = "Top,left,Bottom"
		$tvTreView.add_AfterSelect({
				$Script:lfFolderID = $this.SelectedNode.tag
				Get-ClientFolderItem
				
			})
		$Script:form.Controls.Add($tvTreView)
		
		# Add DataGrid View
		$dgDataGrid = new-object System.windows.forms.DataGridView
		$dgDataGrid.Location = new-object System.Drawing.Size(250, 200)
		$dgDataGrid.size = new-object System.Drawing.Size(800, 600)
		$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
		$dgDataGrid.AllowUserToDeleteRows = $false
		$dgDataGrid.AllowUserToAddRows = $false
		$Script:form.Controls.Add($dgDataGrid)
		
		$Script:form.Text = "Simple Exchange Mailbox Client"
		$Script:form.size = new-object System.Drawing.Size(1200, 800)
		$Script:form.autoscroll = $true
		$Script:form.Add_Shown({ $Script:form.Activate() })
		if ($AccessToken -ne $null)
		{
			$Script:AccessToken = $AccessToken
			$emEmailAddressTextBox.Text = $MailboxName
			Open-ClientMailbox -AccessToken $AccessToken
		}
		$Script:form.ShowDialog()
	}
}