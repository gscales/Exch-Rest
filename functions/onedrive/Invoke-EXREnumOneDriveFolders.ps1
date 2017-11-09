function Invoke-EXREnumOneDriveFolders
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
		[String]
		$DriveRESTURI,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$FolderPath
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		if ([String]::IsNullOrEmpty($DriveRESTURI))
		{
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "('$MailboxName')/drive/root/children?`$filter folder ne null`&`$Top=1000"
		}
		else
		{
			$RequestURL = $DriveRESTURI + "/children?`$filter folder ne null`&`$Top=1000"
		}
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Item in $JSONOutput.value)
			{
				Add-Member -InputObject $Item -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $Item.Id + "')")
				Add-Member -InputObject $Item -NotePropertyName Path -NotePropertyValue ("\" + $Item.name)
				if ([bool]($Item.PSobject.Properties.name -match "folder"))
				{
					write-output $Item
					if ($Item.folder.childCount -gt 0)
					{
						Invoke-EXREnumChildFolders -DriveRESTURI $Item.DriveRESTURI -MailboxName $MailboxName -AccessToken $AccessToken -Path $Item.Path
					}
				}
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
	}
}
