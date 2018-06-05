function Invoke-UpdateItem {
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$details
	)
	Begin {
		if ($AccessToken -eq $null) {
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}

		$HttpClient = Get-HTTPClient($MailboxName)
		$RequestURL = $ItemURI
		$results = Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $details
		return $results		
	}
}