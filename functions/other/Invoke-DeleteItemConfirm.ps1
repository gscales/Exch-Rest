function Invoke-DeleteItem {
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
		$confirmation
	)
	Begin {
		if ($AccessToken -eq $null) {
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		if ($confirmation -ne 'y') {
			$confirmation = Read-Host "Are you Sure You Want To proceed with deleting the Item"
		}
		if ($confirmation -eq 'y') {
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $ItemURI
			$results = & Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			return $results
		}
		else {
			Write-Host "skipped deletion"
		}
	}
}