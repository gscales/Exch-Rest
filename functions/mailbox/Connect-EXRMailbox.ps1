function Connect-EXRMailbox
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$ClientId,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$redirectUrl,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$ClientSecret,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$ResourceURL,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$Beta,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[String]
		$Prompt,

		[Parameter(Position = 7, Mandatory = $false)]
		[switch]
		$SaveToPrivateData,

		[Parameter(Position = 8, Mandatory = $false)]
		[switch]
		$Outlook
		
	)
	Begin
	{
		if ([String]::IsNullOrEmpty($ClientId))
		{
			$Resource = "graph.microsoft.com"
			if($Outlook.IsPresent){
				$Resource = ""
			}
			if($beta.IsPresent){
				return  Get-EXRAccessToken -MailboxName $MailboxName -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl urn:ietf:wg:oauth:2.0:oob -ResourceURL $Resource -beta -Prompt $Prompt -SaveToPrivateData                  
			}
			else{
				return  Get-EXRAccessToken -MailboxName $MailboxName -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl urn:ietf:wg:oauth:2.0:oob -ResourceURL $Resource -Prompt $Prompt -SaveToPrivateData     
			}
		}
		else{
			return Get-AccessToken -ClientId $ClientId -MailboxName $MailboxName -redirectUrl $redirectUrl -ClientSecret $ClientSecret -ResourceURL $ResourceURL -Beta:$beta.IsPresent -prompt $Prompt -SaveToPrivateData:$SaveToPrivateData.isPresent
		}
	}
}
