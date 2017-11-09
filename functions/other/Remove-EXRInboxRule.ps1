function Remove-EXRInboxRule{
    <#
    .SYNOPSIS
    Remove an inbox rule.
    
    .DESCRIPTION
    Remove an inbox rule.
    
    .PARAMETER MailboxName
    The mailbox to query.
    
    .PARAMETER AccessToken
    The access token used to connect to the mailbox.
    
    .PARAMETER Id
    The id of the inbox rule to query.
    
    .EXAMPLE
    Remove inbox rule with id 'AgAAAFYTrjs='
    Remove-EXRInboxRule $MailboxName $AccessToken "AgAAAFYTrjs="
    
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true)]
        [string]$MailboxName,
        
        [Parameter(Position=1, Mandatory=$false)]
        [psobject]$AccessToken,
        
        [Parameter(Position=2, Mandatory=$true)]
        [string]$Id
    )
    Begin{
        if($AccessToken -eq $null){
              $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName -Beta
        }
        elseif(!$AccessToken.Beta){
            Throw("This function requires a beta access token. Use the '-Beta' switch with Get-EXRAccessToken to create a beta access token.")
        }
        
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox/MessageRules/$Id"
    }
    Process{
        Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}
