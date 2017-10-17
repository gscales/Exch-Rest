function Get-DefaultContactsFolder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
   
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users" 
        $RequestURL =   $EndPoint + "('$MailboxName')/contactFolders/contacts"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

