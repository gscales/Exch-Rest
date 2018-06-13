function Get-EXRAuditLogsSignIns {
    [CmdletBinding()]
    param( 
        [Parameter(Position = 0, Mandatory = $false)] [string]$filter,
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [psobject]$MailboxName,
        [Parameter(Position = 2, Mandatory = $false)] [String]$Id
        
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
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "auditLogs" -beta
        $RequestURL = $EndPoint + "/signIns?Top=100"
        if (![String]::IsNullOrEmpty($Id)) {
            $RequestURL += "/" + $Id
        }
        if (![String]::IsNullOrEmpty($filter)) {
            $RequestURL += "&`$filter=" + $filter
        }   
        do {
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            Write-Output $JSONOutput.value   
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while (![String]::IsNullOrEmpty($RequestURL))

     
        
        
    }
}
