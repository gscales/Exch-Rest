function Get-AllMailFolders
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[PSCustomObject]
		$PropList
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/msgfolderroot/childfolders?`$Top=1000"
		if ($PropList -ne $null)
		{
			$Props = Get-ExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
			$RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
			Write-Host $RequestURL
		}
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Folder in $JSONOutput.Value)
			{
				$Folder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ("\\" + $Folder.DisplayName)
				$folderId = $Folder.Id.ToString()
				Add-Member -InputObject $Folder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
				Write-Output $Folder
				if ($Folder.ChildFolderCount -gt 0)
				{
					if ($PropList -ne $null)
					{
						Get-AllChildFolders -Folder $Folder -AccessToken $AccessToken -PropList $PropList
					}
					else
					{
						Get-AllChildFolders -Folder $Folder -AccessToken $AccessToken
					}
				}
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}
