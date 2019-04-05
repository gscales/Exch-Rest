function Get-EXRTeamsPrivateChatMessages{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position = 3, Mandatory = $false)] [DateTime]$startdatetime = (Get-Date).AddDays(-365),
        [Parameter(Position = 4, Mandatory = $false)] [datetime]$enddatetime = (Get-Date),
        [Parameter(Position = 5, Mandatory = $false)] [String]$SenderAddress
    )
    Begin{
        $rptCollection = @()
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
        $Filter = "receivedDateTime ge " + $startdatetime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") + " and receivedDateTime le " + $enddatetime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $Filter += " and singleValueExtendedProperties/any(ep: ep/id eq 'String 0x001a' and ep/value eq 'IPM.SkypeTeams.Message')" 
        if(![String]::IsNullOrEmpty($SenderAddress)){
            $Filter += " and  from/emailAddress/address eq '" + $SenderAddress + "'"
        }
                
        $Items = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder AllItems -Filter $Filter -AdditionalProperties "lastModifiedDateTime"
        return $Items
        
    }
}