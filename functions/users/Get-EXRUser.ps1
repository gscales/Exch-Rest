function Get-EXRUser {
    [CmdletBinding()]
    param( 
        [Parameter(Position = 0, Mandatory = $false)] [string]$UPN,
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$SelectProperties
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
        if ([String]::IsNullOrEmpty($UPN)) {
            $UPN = $MailboxName
        }
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "('" + [uri]::EscapeDataString($UPN) + "')"
        if(![String]::IsNullorEmpty($SelectProperties)){            
            $RequestURL += "?`$select=id,mail,displayName,userPrincipalName,givenName,surname," + $SelectProperties
        }
        $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        return $JSONOutput
      
     
        
        
    }
}
