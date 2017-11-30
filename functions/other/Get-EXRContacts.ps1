function Get-EXRContacts{

	   [CmdletBinding()] 
    param( 
		[Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
	    [Parameter(Position=2, Mandatory=$false)] [string]$MailboxName,
		[Parameter(Position=3, Mandatory=$false)] [string]$ContactsFolderName
		

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
			if([String]::IsNullOrEmpty($ContactsFolderName)){
			    $Contacts = Get-EXRDefaultContactsFolder -MailboxName $MailboxName -AccessToken $AccessToken
			}
			else{
			    $Contacts = Get-EXRContactsFolder -MailboxName $MailboxName -AccessToken $AccessToken -FolderName $ContactsFolderName
			    if([String]::IsNullOrEmpty($Contacts)){throw "Error Contacts folder not found check the folder name this is case sensitive"}
			}    
            $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users" 
            $RequestURL =  $EndPoint + "('" + $MailboxName + "')/contactFolders('" + $Contacts.id  + "')/contacts/?`$Top=1000"
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Message in $JSONOutput.Value) {
					Write-Output $Message
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL)) 

	} 
}

