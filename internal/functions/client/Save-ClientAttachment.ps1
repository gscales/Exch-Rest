function Save-ClientAttachment
{
	[CmdletBinding()]
	Param (
		
	)
	
	$dlfolder = new-object -ComObject shell.application
	$dlfolderpath = $dlfolder.BrowseForFolder(0, "Download attachments to", 0)
	Get-EXRAttachments -MailboxName $emEmailAddressTextBox.Text -ItemURI $Script:msMessage.ItemRESTURI -MetaData -AccessToken $Script:AccessToken | ForEach-Object{
		$attach = Invoke-EXRDownloadAttachment -MailboxName $emEmailAddressTextBox.Text -AttachmentURI $_.AttachmentRESTURI -AccessToken $Script:AccessToken
		$fiFile = new-object System.IO.FileStream(($dlfolderpath.Self.Path + "\" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
		$attachBytes = [System.Convert]::FromBase64String($attach.ContentBytes)
		$fiFile.Write($attachBytes, 0, $attachBytes.Length)
		$fiFile.Close()
		write-host ("Downloaded Attachment : " + (($dlfolderpath.Self.Path + "\" + $attach.Name.ToString())))
	}
}