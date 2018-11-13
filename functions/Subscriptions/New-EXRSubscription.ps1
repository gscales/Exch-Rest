function New-EXRSubscription
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
        $AccessToken,
        
        [Parameter(Position = 2, Mandatory = $false)]
		[string]
        $changeType,
        
        [Parameter(Position = 3, Mandatory = $false)]
		[string]
        $notificationUrl,
        
        [Parameter(Position = 4, Mandatory = $false)]
		[string]
        $resource,
        
        [Parameter(Position = 5, Mandatory = $false)]
		[datetime]
        $expirationDateTime,
        
        [Parameter(Position = 6, Mandatory = $false)]
		[string]
		$clientState

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
		$EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "Subscriptions"
        $RequestURL = $EndPoint 
        $NewSubscription = @{}
        $NewSubscription.Add("changeType",$changeType)
        $NewSubscription.Add("notificationUrl",$notificationUrl)
        $NewSubscription.Add("resource",$resource)
        $NewSubscription.Add("expirationDateTime",$expirationDateTime)
        $NewSubscription.Add("clientState",$clientState)
        $JSONOutput = Invoke-RestPost -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content (ConvertTo-Json $NewSubscription -Depth 8)
        return $JSONOutput

	}
}
