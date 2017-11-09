function Open-ClientMailbox
{
	[CmdletBinding()]
	param (
		[psobject]
		$AccessToken
	)
	
	Process
	{
		$tvTreView.Nodes.Clear()
		$Script:Treeinfo.Clear()
		if ($AccessToken -eq $null)
		{
			$Script:AccessToken = Get-EXRAccessToken -MailboxName $emEmailAddressTextBox.Text -ClientId $unCASUrlTextBox.Text -redirectUrl $RedirectTextBox.Text -ResourceURL graph.Microsoft.com
		}
		else
		{
			$Script:AccessToken = $AccessToken
		}
		$rootFolder = Get-EXRRootMailFolder -AccessToken $Script:AccessToken -MailboxName $emEmailAddressTextBox.Text
		if ($ShowFolderSize)
		{
			$PropList = @()
			$FolderSizeProp = Get-EXRTaggedProperty -Id "0x0E08" -DataType Long
			$PropList += $FolderSizeProp
			$Folders = Get-EXRAllMailFolders -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken -PropList $PropList
		}
		else
		{
			$Folders = Get-EXRAllMailFolders -MailboxName $emEmailAddressTextBox.Text -AccessToken $Script:AccessToken
		}
		$Script:Treeinfo = @{ }
		$TNRoot = new-object System.Windows.Forms.TreeNode("Root")
		$TNRoot.Name = "Mailbox"
		$TNRoot.Text = "Mailbox - " + $emEmailAddressTextBox.Text
		$exProgress = 0
		foreach ($ffFolder in $Folders)
		{
			#Process folder here
			$ParentFolderId = $ffFolder.parentFolderId
			$folderName = $ffFolder.displayName
			if ($ShowFolderSize)
			{
				$folderName = $ffFolder.displayName + " (" + [math]::round($ffFolder.singleValueExtendedProperties[0].value /1Mb, 0) + " mb)"
			}
			$TNChild = new-object System.Windows.Forms.TreeNode($ffFolder.Name)
			$TNChild.Name = $folderName
			$TNChild.Text = $folderName
			$TNChild.tag = $ffFolder
			
			if ($ParentFolderId -eq $rootFolder.Id)
			{
				[void]$TNRoot.Nodes.Add($TNChild)
				$Script:Treeinfo.Add($ffFolder.Id.ToString(), $TNChild)
			}
			else
			{
				$pfFolder = $Script:Treeinfo[$ParentFolderId]
				[void]$pfFolder.Nodes.Add($TNChild)
				if ($Script:Treeinfo.ContainsKey($ffFolder.Id) -eq $false)
				{
					$Script:Treeinfo.Add($ffFolder.Id, $TNChild)
				}
			}
		}
		$Script:clickedFolder = $null
		[void]$tvTreView.Nodes.Add($TNRoot)
		Write-Progress -Activity "Executing Request" -Completed
	}
}