function Get-EXRInboxRule{
    <#
    .SYNOPSIS
    Get inbox rules.
    
    .DESCRIPTION
    Get inbox rules.
    
    .PARAMETER MailboxName
    The mailbox to query.
    
    .PARAMETER AccessToken
    The access token used to connect to the mailbox.
    
    .PARAMETER Id
    Optional. The id of the inbox rule to query.
    
    .PARAMETER DisplayName
    Optional. Get display name to against.
    
    .EXAMPLE
    Get all inbox rules
    Get-EXRInboxRule "username@example.com" $AccessToken
    
    .EXAMPLE
    Get inbox rule with id 'AgAAAALK7kQ='
    Get-EXRInboxRule "username@example.com" $AccessToken "AgAAAALK7kQ="
    
    #>
    [CmdletBinding(DefaultParameterSetName = '__AllParameterSets')]
    param(
        [Parameter(Position=0, Mandatory=$false)]
        [string]$MailboxName,
        
        [Parameter(Position=1, Mandatory=$False)]
        [psobject]$AccessToken,
        
        [Parameter(Position=2, Mandatory=$False, ParameterSetName='Id')]
        [string]$Id,
        
        [Parameter(Position=3, Mandatory=$False, ParameterSetName='DisplayName')]
        [psobject]$DisplayName
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
         
        $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users" -beta
        if($PSCmdLet.ParameterSetName -eq "Id"){
            $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox/MessageRules/$Id"
        }
        else{
            $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox/MessageRules"
        }
    }
    Process{
        $Results = @()
        $Response = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        
        if($PSCmdLet.ParameterSetName -eq "Id")
        {
            [void]$Response.PSObject.TypeNames.Insert(0, "PoshExchRest.InboxRule")
            $Results += $Response
        }
        else{
            foreach($Entry in $Response.Value){
                [void]$Entry.PSObject.TypeNames.Insert(0, "PoshExchRest.InboxRule")
                if($PSCmdLet.ParameterSetName -eq "DisplayName"){
                    if($Entry.DisplayName -eq $DisplayName){
                        $Results += $Entry
                    }
                }
                else{
                    $Results += $Entry
                }
            }
        }
        
        return $Results
    }
}
