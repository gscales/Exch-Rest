function New-EXRFolder
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$ParentFolderPath,

		[Parameter(Position = 3, Mandatory = $false)]
		[switch]
		$RootFolder,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[string]
		$DisplayName,

		[Parameter(Position = 5, Mandatory = $false)]
		[string]
		$FolderClass
		
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
		if($RootFolder.IsPresent){
			$ParentFolder = Get-EXRRootMailFolder -AccessToken $AccessToken -MailboxName $MailboxName
		}else{
			$ParentFolder = Get-EXRFolderFromPath -FolderPath $ParentFolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		}
		
		if ($ParentFolder -ne $null)
		{
			$HttpClient = Get-HTTPClient -MailboxName $MailboxName
			$RequestURL = $ParentFolder.FolderRestURI + "/childfolders"
			if([String]::IsNullOrEmpty($FolderClass)){
				$NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"}"
			}
			else{
				$NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"," +"`r`n"
				$NewFolderPost += "`"SingleValueExtendedProperties`": [{`"PropertyId`":`"String 0x3613`",`"Value`":`"" + $FolderClass + "`"}]}"
			}
			
			write-host $NewFolderPost
			return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
			
		}
		
		
	}
}
