function Set-EXRReadFlag
{
	[CmdletBinding()]
	param (

	    [parameter(ValueFromPipeline=$True)]
		[psobject[]]$Item,
		
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,

		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$ItemId,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
		$AccessToken		
	)
	Process
	{
		
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
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('" + $MailboxName + "')/MailFolders/AllItems/messages"
		if($Item -eq $null){
			$RequestURL =  $EndPoint + "('" + $MailboxName + "')/MailFolders/AllItems/messages" + "('" + $ItemId + "')"
		}
		else{
			$RequestURL =  $EndPoint + "('" + $MailboxName + "')/MailFolders/AllItems/messages" + "('" + $Item.Id + "')"
		}
        $UpdateProps = @()
        $UpdateProps += (Get-EXRItemProp -Name IsRead -Value true -NoQuotes)
		$UpdateItemPatch = Get-MessageJSONFormat -StandardPropList $UpdateProps
		return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $UpdateItemPatch
	}
}
