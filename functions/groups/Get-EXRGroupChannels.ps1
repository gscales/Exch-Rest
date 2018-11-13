function Get-EXRGroupChannels{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [psobject]$Group,
        [Parameter(Position=3, Mandatory=$false)] [psobject]$ChannelName
    )
    Process{
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
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "groups"
        if([String]::IsNullOrEmpty($ChannelName)){
            $RequestURL =   $EndPoint + "('" + $Group.Id + "')/channels?`$Top=1000"
        }
        else{
            $RequestURL =   $EndPoint + "('" + $Group.Id + "')/channels?`$filter=displayName eq '$ChannelName'"
        }
        Write-Host $RequestURL        
        do{
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                Write-Output $Message
            }           
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while(![String]::IsNullOrEmpty($RequestURL))       
        
        
    }
}
