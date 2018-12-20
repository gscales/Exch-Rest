function Invoke-EXRDeleteItem {
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$confirmation
	)
	Begin {
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
		if ($confirmation -ne 'y') {
			$confirmation = Read-Host "Are you Sure You Want To proceed with deleting the Item (Y/N)"
		}
		if ($confirmation.tolower() -eq 'y') {
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