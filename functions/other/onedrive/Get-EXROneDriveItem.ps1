function Get-EXROneDriveItem
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
		[String]
		$DriveRESTURI
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
		$RequestURL = $DriveRESTURI
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Add-Member -InputObject $JSONOutput -NotePropertyName DriveRESTURI -NotePropertyValue (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $JSONOutput.Id + "')")
		return $JSONOutput
	}
}
