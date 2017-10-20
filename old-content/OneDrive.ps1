
#region OneDrive
function Get-DefaultOneDrive
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
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/drive/root"
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Add-Member -InputObject $JSONOutput -NotePropertyName DriveRESTURI -NotePropertyValue ($EndPoint + "('$MailboxName')/drives('" + $JSONOutput.Id + "')")
		return $JSONOutput
		
		
	}
}

function Get-DefaultOneDriveRootItems
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
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/drive/root/children"
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		foreach ($Item in $JSONOutput.value)
		{
			Add-Member -InputObject $Item -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $Item.Id + "')")
			write-output $Item
		}
		
		
		
	}
}

function Invoke-EnumOneDriveFolders
{
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
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
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
						Invoke-EnumChildFolders -DriveRESTURI $Item.DriveRESTURI -MailboxName $MailboxName -AccessToken $AccessToken -Path $Item.Path
					}
				}
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
	}
}

function Invoke-EnumChildFolders
{
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
		$Path
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		if ([String]::IsNullOrEmpty($DriveRESTURI))
		{
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $FolderPath + ":/children?`$filter folder ne null`&`$Top=1000"
		}
		else
		{
			$RequestURL = $DriveRESTURI + "/children?`$filter folder ne null`&`$Top=1000"
		}
		$pc = 0;
		do
		{
			$pc++
			#write-host "Page " + $pc
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Item in $JSONOutput.value)
			{
				Add-Member -InputObject $Item -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $Item.Id + "')")
				Add-Member -InputObject $Item -NotePropertyName Path -NotePropertyValue ($Path + "\" + $Item.name)
				if ([bool]($Item.PSobject.Properties.name -match "folder"))
				{
					write-output $Item
					if ($Item.folder.childCount -gt 0)
					{
						Invoke-EnumChildFolders -DriveRESTURI $Item.DriveRESTURI -MailboxName $MailboxName -AccessToken $AccessToken -Path $Item.Path
					}
				}
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
		
	}
}

function Invoke-FolderPicker
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$rootFolder,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psObject]
		$Folders,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$pickerType,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[switch]
		$ShowFolderSize
	)
	Begin
	{
		
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
		$Treeinfo = @{ }
		$TNRoot = new-object System.Windows.Forms.TreeNode("Root")
		$TNRoot.Name = "Mailbox"
		$TNRoot.Text = "Mailbox - " + $MailboxName
		foreach ($ffFolder in $Folders)
		{
			#Process folder here
			switch ($pickerType)
			{
				"onedrive" {
					$ParentFolderId = $ffFolder.parentReference.Id
					$folderName = $ffFolder.Name
				}
				"mail" {
					$ParentFolderId = $ffFolder.parentFolderId
					$folderName = $ffFolder.displayName
					if ($ShowFolderSize)
					{
						$folderName = $ffFolder.displayName + " (" + [math]::round($ffFolder.singleValueExtendedProperties[0].value /1Mb, 0) + " mb)"
					}
				}
			}
			$TNChild = new-object System.Windows.Forms.TreeNode($ffFolder.Name)
			$TNChild.Name = $folderName
			$TNChild.Text = $folderName
			$TNChild.tag = $ffFolder
			
			
			if ($ParentFolderId -eq $rootFolder.Id)
			{
				[void]$TNRoot.Nodes.Add($TNChild)
				$Treeinfo.Add($ffFolder.Id.ToString(), $TNChild)
			}
			else
			{
				$pfFolder = $Treeinfo[$ParentFolderId]
				[void]$pfFolder.Nodes.Add($TNChild)
				if ($Treeinfo.ContainsKey($ffFolder.Id) -eq $false)
				{
					$Treeinfo.Add($ffFolder.Id, $TNChild)
				}
			}
		}
		$Script:clickedFolder = $null
		$objForm = New-Object System.Windows.Forms.Form
		$objForm.Text = "Folder Select Form"
		$objForm.Size = New-Object System.Drawing.Size(600, 600)
		$objForm.StartPosition = "CenterScreen"
		$tvTreView1 = new-object System.Windows.Forms.TreeView
		$tvTreView1.Location = new-object System.Drawing.Size(1, 1)
		$tvTreView1.add_DoubleClick({
				$Script:clickedFolder = $this.SelectedNode.tag
				$objForm.Close()
			})
		$tvTreView1.size = new-object System.Drawing.Size(580, 580)
		$tvTreView1.Anchor = "Top,left,Bottom"
		[void]$tvTreView1.Nodes.Add($TNRoot)
		$objForm.controls.add($tvTreView1)
		[void]$objForm.ShowDialog()
		return, $Script:clickedFolder
		
	}
}

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

function Get-OneDriveChildren
{
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
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		if ([String]::IsNullOrEmpty($DriveRESTURI))
		{
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $FolderPath + ":/children"
		}
		else
		{
			$RequestURL = $DriveRESTURI + "/children"
		}
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		foreach ($Item in $JSONOutput.value)
		{
			Add-Member -InputObject $Item -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $Item.Id + "')")
			write-output $Item
		}
		
	}
}

function Get-OneDriveItem
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$DriveRESTURI
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$RequestURL = $DriveRESTURI
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Add-Member -InputObject $JSONOutput -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $JSONOutput.Id + "')")
		return $JSONOutput
	}
}

function Get-OneDriveItemFromPath
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$OneDriveFilePath
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $OneDriveFilePath
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Add-Member -InputObject $JSONOutput -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $JSONOutput.Id + "')")
		return $JSONOutput
	}
}

function Invoke-UploadOneDriveItemToPath
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$OneDriveUploadFilePath,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$FilePath,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[Byte[]]
		$FileBytes
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $OneDriveUploadFilePath + ":/content"
		if ([String]::IsNullOrEmpty($FileBytes))
		{
			$Content = ([System.IO.File]::ReadAllBytes($filePath))
		}
		else
		{
			$Content = $FileBytes
		}
		$JSONOutput = Invoke-RestPut -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -content $Content -contentheader "application/octet-stream"
		return $JSONOutput
	}
}

function New-ReferanceAttachment
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$SourceUrl,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$ProviderType,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[String]
		$Permission,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$IsFolder
		
	)
	Begin
	{
		$ReferanceAttachment = "" | Select-Object Name, SourceUrl, ProviderType, Permission, IsFolder
		$ReferanceAttachment.IsFolder = "False"
		$ReferanceAttachment.ProviderType = "oneDriveBusiness"
		$ReferanceAttachment.Permission = $Permission
		$ReferanceAttachment.SourceUrl = $SourceUrl
		$ReferanceAttachment.Name = $Name
		if (![String]::IsNullOrEmpty($ProviderType))
		{
			$ReferanceAttachment.ProviderType = $ProviderType
		}
		if (![String]::IsNullOrEmpty($IsFolder))
		{
			$ReferanceAttachment.IsFolder = $IsFolder
		}
		return $ReferanceAttachment
	}
}
#endregion