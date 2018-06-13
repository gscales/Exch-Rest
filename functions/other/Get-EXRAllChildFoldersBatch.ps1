function Get-EXRAllChildFoldersBatch
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[PSCustomObject]
		$BatchItems,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[PSCustomObject]
		$PropList,

		[Parameter(Position = 4, Mandatory = $false)]
		[PSCustomObject]
		$MailboxName,

		
		[Parameter(Position = 5, Mandatory = $false)]
		[PSCustomObject]
		$FldIndex

	)
	Begin
	{
		$ChildFolders = Get-EXRBatchItems -Items $BatchItems -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList -URLString ("/users" + "('" + $MailboxName + "')" + "/MailFolders") -ChildFolders		
		$BatchItems = @()
		for($intcnt=0;$intcnt -lt $ChildFolders.value.Count;$intcnt++){
				$Child = $ChildFolders.value[$intcnt]					
				if($Child -ne $null){
					$ParentId = $Child.parentFolderId
					$ChildFldPath = ($FldIndex[$ParentId] + "\" + $Child.displayName)
					$Child | Add-Member -NotePropertyName FolderPath -NotePropertyValue $ChildFldPath
					$FldIndex.Add($Child.Id,$ChildFldPath)
					$folderId = $Child.Id.ToString()
					Add-Member -InputObject $Child -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
					Expand-ExtendedProperties -Item $Child
					Write-Output $Child
					if ($Child.ChildFolderCount -gt 0)
					{
						$BatchItems += $Child
					}
					if($BatchItems.Count -gt 20){
						Get-EXRAllChildFoldersBatch -BatchItems $BatchItems -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList -FldIndex $FldIndex
						$BatchItems = @()
					}	
				}
		}
		if($BatchItems.Count -gt 0){
			Get-EXRAllChildFoldersBatch -BatchItems $BatchItems -MailboxName $MailboxName -AccessToken $AccessToken -PropList $PropList -FldIndex $FldIndex	
			$BatchItems = @()
		}
		
	}
}
