function Invoke-EXROneDriveFolderPicker
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
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
		$rootFolder = Get-EXRDefaultOneDrive -AccessToken $AccessToken -MailboxName $MailboxName
		$Folders = Invoke-EXREnumOneDriveFolders -MailboxName $MailboxName -AccessToken $AccessToken
		Invoke-EXRFolderPicker -MailboxName $MailboxName -Folders $Folders -rootFolder $rootFolder -pickerType onedrive
	}
}
