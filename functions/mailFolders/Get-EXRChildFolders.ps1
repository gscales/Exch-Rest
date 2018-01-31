function Get-EXRChildFolders
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[PSCustomObject]
		$Folder,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[PSCustomObject]
		$PropList,

		[Parameter(Position = 3, Mandatory = $false)]
		[PSCustomObject]
		$MailboxName,

		[Parameter(Position = 4, Mandatory = $false)]
		[switch]
		$Recurse
	)
	Begin
	{
		if($AccessToken -eq $null)
		{
			$AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
			if($AccessToken -eq $null){
				$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
			}                 
		}
		if([String]::IsNullOrEmpty($MailboxName)){
			$MailboxName = $AccessToken.mailbox
		}  
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $Folder.FolderRestURI + "/childfolders/?`$Top=1000"
		if ($PropList -ne $null)
		{
			$Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
			$RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
		}
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -TrackStatus $true -ProcessMessage ("Processing Folder " + $Folder.FolderPath )
			foreach ($ChildFolder in $JSONOutput.Value)
			{
				$ChildFolder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ($Folder.FolderPath + "\" + $ChildFolder.DisplayName)
				$folderId = $ChildFolder.Id.ToString()
				Add-Member -InputObject $ChildFolder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
				Expand-ExtendedProperties -Item $ChildFolder
				Write-Output $ChildFolder
				if($Recurse.IsPresent){
					if ($ChildFolder.ChildFolderCount -gt 0)
					{
						if ($PropList -ne $null)
						{
							Get-EXRAllChildFolders -Folder $ChildFolder -AccessToken $AccessToken -PropList $PropList -MailboxName $MailboxName
						}
						else
						{
							Get-EXRAllChildFolders -Folder $ChildFolder -AccessToken $AccessToken -MailboxName $MailboxName
						}
					}
				}
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}
