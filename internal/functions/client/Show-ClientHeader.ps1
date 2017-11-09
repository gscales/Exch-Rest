function Show-ClientHeader
{
	[CmdletBinding()]
	Param (
		
	)
	
	$MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$script:msMessage = Get-EXREmail -MailboxName $emEmailAddressTextBox.Text -ItemRESTURI $MessageID -AccessToken $Script:AccessToken -PropList (Get-EXRTransportHeader)
	write-host $MessageID
	$hdrform = new-object System.Windows.Forms.form
	$hdrform.Text = $script:msMessage.Subject
	$hdrform.size = new-object System.Drawing.Size(800, 600)
	
	# Add Message header
	$miMessageHeadertextlabelBox = new-object System.Windows.Forms.RichTextBox
	$miMessageHeadertextlabelBox.Location = new-object System.Drawing.Size(10, 10)
	$miMessageHeadertextlabelBox.size = new-object System.Drawing.Size(800, 600)
	$miMessageHeadertextlabelBox.text = $script:msMessage.PR_TRANSPORT_MESSAGE_HEADERS
	$hdrform.controls.Add($miMessageHeadertextlabelBox)
	$hdrform.autoscroll = $true
	$hdrform.Add_Shown({ $Script:form.Activate() })
	$hdrform.ShowDialog()
}