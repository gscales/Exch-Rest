function Invoke-EXRDeleteContact {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 4, Mandatory = $true)] [string]$id,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Confirm
    )  
    Begin {
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        }  
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users" 
        $RequestURL = $EndPoint + "('" + $MailboxName + "')/Contacts('" + $id + "')"
        if($Confirm.IsPresent){
            $confirmation = "y"
        }
        if ($confirmation -ne 'y') {
			$confirmation = Read-Host "Are you Sure You Want To proceed with deleting the Item"
		}
		if ($confirmation -eq 'y') {		
			$results = & Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			return $results
		}
		

    } 
}

