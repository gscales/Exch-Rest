function New-EXRContactFolder
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$DisplayName,

		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$ParentFolderName
		
		
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
		if(![String]::IsNullOrEmpty($ParentFolderName)){
			$prf = Get-EXRContactsFolder -FolderName $ParentFolderName -MailboxName $MailboxName
			$RequestURL = $EndPoint + "('$MailboxName')/ContactFolders('" + $prf.id + "')/childFolders"
		}
		else{
			$RequestURL = $EndPoint + "('$MailboxName')/ContactFolders"
		}
		
		$NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"}"
		write-host $NewFolderPost
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
		
		
	}
}
