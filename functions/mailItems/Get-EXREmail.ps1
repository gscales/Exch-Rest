function Get-EXREmail
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
		[psobject]
		$ItemRESTURI,

		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
		$PropList
	)
	Process
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
		$RequestURL = $ItemRESTURI
		if($PropList -ne $null){
               $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
               $RequestURL += "?`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        }
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Add-Member -InputObject $JSONOutput -NotePropertyName ItemRESTURI -NotePropertyValue $ItemRESTURI
		Expand-ExtendedProperties -Item $JSONOutput
		return $JSONOutput 
		
	}
}
