function Invoke-EXRFolderPicker
{
	[CmdletBinding()]
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
