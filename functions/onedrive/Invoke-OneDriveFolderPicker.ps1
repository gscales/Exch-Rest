function Invoke-OneDriveFolderPicker
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
		$rootFolder = Get-DefaultOneDrive -AccessToken $AccessToken -MailboxName $MailboxName
		$Folders = Invoke-EnumOneDriveFolders -MailboxName $MailboxName -AccessToken $AccessToken
		Invoke-FolderPicker -MailboxName $MailboxName -Folders $Folders -rootFolder $rootFolder -pickerType onedrive
	}
}
