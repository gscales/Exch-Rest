function Get-ClientFolderItem
{
	[CmdletBinding()]
	Param (
		
	)
	
	$mbtable.Clear()
	$folder = $Script:lfFolderID
	if ($seSearchCheck.Checked)
	{
		switch ($snSearchPropDrop.SelectedItem.ToString())
		{
			"Subject" {
				$sfilter = "Subject eq '" + $sbSearchTextBox.Text.ToString() + "'"
				$Items = Get-EXRFolderItems -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ReturnSize -Folder $folder -TopOnly:$true -Top 100 -Filter $sfilter -TrackStatus
			}
			"Body" {
				$sfilter = "`"Body:'" + $sbSearchTextBox.Text.ToString() + "'`""
				$Items = Get-EXRFolderItems -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ReturnSize -Folder $folder -TopOnly:$true -Top 100 -Search $sfilter -TrackStatus
			}
			"From" {
				$sfilter = "`"From:'" + $sbSearchTextBox.Text.ToString() + "'`""
				$Items = Get-EXRFolderItems -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ReturnSize -Folder $folder -TopOnly:$true -Top 100 -Search $sfilter -TrackStatus
			}
		}
	}
	else
	{
		$Items = Get-EXRFolderItems -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -ReturnSize -Folder $folder -TopOnly:$true -Top 100 -TrackStatus
	}
	foreach ($mail in $Items)
	{
		if ($mail.sender.emailAddress.name -ne $null) { $fnFromName = $mail.sender.emailAddress.name }
		else { $fnFromName = "N/A" }
		if ($mail.Subject -ne $null) { $sbSubject = $mail.Subject.ToString() }
		else { $sbSubject = "N/A" }
		if ([bool]($mail.PSobject.Properties.name -match "Size"))
		{
			$mbtable.rows.add($fnFromName, $sbSubject, $mail.receivedDateTime, $mail.Size.ToString(), $mail.ItemRESTURI, $mail.hasAttachments)
		}
		else
		{
			$mbtable.rows.add($fnFromName, $sbSubject, $mail.receivedDateTime, 0, $mail.ItemRESTURI, $mail.hasAttachments)
		}
	}
	$dgDataGrid.DataSource = $mbtable
}