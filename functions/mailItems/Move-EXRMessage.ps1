function Move-EXRMessage
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$TargetFolderPath,

		[Parameter(Position = 4, Mandatory = $false)]
		[psobject]
		$Folder
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
		if (![String]::IsNullOrEmpty($TargetFolderPath))
		{
			$Folder = Get-EXRFolderFromPath -FolderPath $TargetFolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		}
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient -MailboxName $MailboxName
			$RequestURL = $ItemURI + "/move"
			$MoveItemPost = "{`"DestinationId`": `"" + $Folder.Id + "`"}"
			write-host $MoveItemPost
			return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $MoveItemPost
		}
	}
}
