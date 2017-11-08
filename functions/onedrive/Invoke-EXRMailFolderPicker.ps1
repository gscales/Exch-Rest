function Invoke-EXRMailFolderPicker
{
	[CmdletBinding()]
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
		
		$rootFolder = Get-EXRRootMailFolder -AccessToken $AccessToken -MailboxName $MailboxName
		if ($ShowFolderSize)
		{
			$PropList = @()
			$FolderSizeProp = Get-EXRTaggedProperty -Id "0x0E08" -DataType Long
			$PropList += $FolderSizeProp
			$Folders = Get-EXRAllMailFolders -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList
		}
		else
		{
			$Folders = Get-EXRAllMailFolders -MailboxName $MailboxName -AccessToken $AccessToken
		}
		
		
		if ($ShowFolderSize)
		{
			Invoke-EXRFolderPicker -MailboxName $MailboxName -Folders $Folders -rootFolder $rootFolder -pickerType mail -ShowFolderSize
		}
		else
		{
			Invoke-EXRFolderPicker -MailboxName $MailboxName -Folders $Folders -rootFolder $rootFolder -pickerType mail
		}
	}
}
