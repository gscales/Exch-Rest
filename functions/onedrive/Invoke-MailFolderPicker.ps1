function Invoke-MailFolderPicker
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$ShowFolderSize
		
	)
	Begin
	{
		
		$rootFolder = Get-RootMailFolder -AccessToken $AccessToken -MailboxName $MailboxName
		if ($ShowFolderSize)
		{
			$PropList = @()
			$FolderSizeProp = Get-TaggedProperty -Id "0x0E08" -DataType Long
			$PropList += $FolderSizeProp
			$Folders = Get-AllMailFolders -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList
		}
		else
		{
			$Folders = Get-AllMailFolders -MailboxName $MailboxName -AccessToken $AccessToken
		}
		
		
		if ($ShowFolderSize)
		{
			Invoke-FolderPicker -MailboxName $MailboxName -Folders $Folders -rootFolder $rootFolder -pickerType mail -ShowFolderSize
		}
		else
		{
			Invoke-FolderPicker -MailboxName $MailboxName -Folders $Folders -rootFolder $rootFolder -pickerType mail
		}
	}
}
