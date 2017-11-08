function Get-EXREmail
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
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
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-EXRHTTPClient -MailboxName $MailboxName
		$RequestURL = $ItemRESTURI
		if($PropList -ne $null){
               $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
               $RequestURL += "?`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
        }
		$JSONOutput = Invoke-EXRRestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Add-Member -InputObject $JSONOutput -NotePropertyName ItemRESTURI -NotePropertyValue $ItemRESTURI
		Invoke-EXRParseExtendedProperties -Item $JSONOutput
		return $JSONOutput 
		
	}
}
