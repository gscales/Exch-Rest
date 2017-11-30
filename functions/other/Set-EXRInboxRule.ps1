function Set-EXRInboxRule{
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
    
    .PARAMETER Rule
    The JSON representation of an inbox rule.
    
    .EXAMPLE
    Update an inbox rule
    $UpdatedRule = @{
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
    $UpdatedRuleJSON = $UpdatedRule | ConvertTo-JSON -Depth 4
    
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$false)]
        [string]$MailboxName,
        
        [Parameter(Position=1, Mandatory=$false)]
        [psobject]$AccessToken,
        
        [Parameter(Position=2, Mandatory=$true)]
        [string]$Id,
        
        [Parameter(Position=3, Mandatory=$true)]
        [psobject]$Rule
    )
    Begin{
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
        if(!$AccessToken.Beta){
            Throw("This function requires a beta access token. Use the '-Beta' switch with Get-EXRAccessToken to create a beta access token.")
        }
        
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox/MessageRules/$Id"
    }
    Process{
        $Result = Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $Rule
        if($Result.Id -eq $Id){
            [void]$Result.PSObject.TypeNames.Insert(0, "PoshExchRest.InboxRule")
            return $Result
        }
        else{
            return $Result
        }
    }
}
