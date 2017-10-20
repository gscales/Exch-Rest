function Get-TestAccessToken{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=0, Mandatory=$false)] [switch]$beta,
        [Parameter(Position=0, Mandatory=$false)] [String]$Prompt

    )
    Begin{
        if($beta.IsPresent){
		    return  Get-AccessToken -MailboxName $MailboxName -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl urn:ietf:wg:oauth:2.0:oob -ResourceURL graph.microsoft.com -beta -Prompt $Prompt                   
        }
        else{
     		return  Get-AccessToken -MailboxName $MailboxName -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl urn:ietf:wg:oauth:2.0:oob -ResourceURL graph.microsoft.com -Prompt $Prompt          
        }
        
    }
}