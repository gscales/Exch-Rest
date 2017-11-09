function Select-ClientAttachment
{
	[CmdletBinding()]
	Param (
		
	)
	
	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
		Multiselect  = $true
	}
	
	[void]$FileBrowser.ShowDialog()
	foreach ($File in $FileBrowser.FileNames)
	{
		$script:Attachments += $File
		$attname += $File + " "
	}
	$miMessageAttachmentslableBox1.Text = $attname
	
}