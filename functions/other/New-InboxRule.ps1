function New-InboxRule{
    <#
    .SYNOPSIS
    Create an inbox rule.
    
    .DESCRIPTION
    Create an inbox rule.
    
    .PARAMETER MailboxName
    The mailbox to query.
    
    .PARAMETER AccessToken
    The access token used to connect to the mailbox.
    
    .PARAMETER Rule
    The JSON representation of an inbox rule.
    
    .EXAMPLE
    Create a new rule
        $NewRule = @{
        DisplayName = "Test"
        Sequence = 99
        IsEnabled = $False
        IsReadOnly = $False
        Conditions = @{
            SubjectContains = @("TEST TEST TEST")
            FromAddresses = @(
                @{
                    EmailAddress = @{
                        Name = "user@example.com"
                        Address = "user@example.com"
                    }
                }
            )
        }
        Actions = @{
            MoveToFolder = "AQMkAGNkZTcwMjllLWU4MjUtNDI0YS1iNWU3LWIxNmFjNDhiM2Y2OQAuAAAD-a164VMCK0K9CBbYzXNdGAEA5ucLn3NEf0aVElSHo0-AfwAAAXgABQAAAA=="
            StopProcessingRules = $True
        }
    }
    $NewRuleJSON = $NewRule | ConvertTo-JSON -Depth 4
    New-InboxRule $MailboxName $AccessToken $NewRuleJSON
    
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$True)]
        [string]$MailboxName,
        
        [Parameter(Position=1, Mandatory=$False)]
        [psobject]$AccessToken,
        
        [Parameter(Position=2, Mandatory=$True)]
        [psobject]$Rule
    )
    Begin{
        if($AccessToken -eq $null){
              $AccessToken = Get-AccessToken -MailboxName $MailboxName -Beta
        }
        elseif(!$AccessToken.Beta){
            Throw("This function requires a beta access token. Use the '-Beta' switch with Get-AccessToken to create a beta access token.")
        }
        
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox/MessageRules"
    }
    Process{
        $Result = Invoke-RestPost -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $Rule
        
        if($Result.Count -eq 2){
            [void]$Result[1].PSObject.TypeNames.Insert(0, "PoshExchRest.InboxRule")
            return $Result[1]
        }
        else{
            return $Result
        }
    }
}
