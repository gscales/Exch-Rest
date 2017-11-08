function Invoke-EXROneDriveFolderPicker
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		$rootFolder = Get-EXRDefaultOneDrive -AccessToken $AccessToken -MailboxName $MailboxName
		$Folders = Invoke-EXREnumOneDriveFolders -MailboxName $MailboxName -AccessToken $AccessToken
		Invoke-EXRFolderPicker -MailboxName $MailboxName -Folders $Folders -rootFolder $rootFolder -pickerType onedrive
	}
}
