function Connect-EXRMailbox
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
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
		$CacheCredentials,

		[Parameter(Position = 8, Mandatory = $false)]
		[switch]
		$Outlook
		
	)
	Begin
	{
		if ([String]::IsNullOrEmpty($ClientId))
		{
			$ProceedOkay = $false
			Do {
				Write-Host "
				---Default ClientId Selection ----------
				1 = Mailbox Access Only
				2 = Mailbox Contacts Access Only
				3 = Full Access to all Graph API functions
				4 = Reporting Access Only
				--------------------------"
				$choice1 = read-host -prompt "Select number & press enter"
				switch($choice1){
					"1"{
						$ProceedOkay = $true
						$ClientId = "1d236c67-7e0b-42bc-88fd-d0b70a3df50a"
					}
					"2" {
						$ProceedOkay = $true
						$ClientId = "9149e700-47a9-4ba6-b01e-20716509fac7"
						
					}
					"3" {
						$ProceedOkay = $true
						$ClientId = "5471030d-f311-4c5d-91ef-74ca885463a7"
					}
					"4" {
						$ProceedOkay = $true
						$ClientId = "e9a8cb7e-9630-4313-8705-9d6f3181bf01"
					}
				}
			} until ($ProceedOkay)
			$Resource = "graph.microsoft.com"
			if($Outlook.IsPresent){
				$Resource = ""
			}
			if($beta.IsPresent){
				return  Get-EXRAccessToken -MailboxName $MailboxName -ClientId $ClientId  -redirectUrl urn:ietf:wg:oauth:2.0:oob -ResourceURL $Resource -beta -Prompt $Prompt -CacheCredentials                  
			}
			else{
				return  Get-EXRAccessToken -MailboxName $MailboxName -ClientId $ClientId -redirectUrl urn:ietf:wg:oauth:2.0:oob -ResourceURL $Resource -Prompt $Prompt -CacheCredentials   
			}
		}
		else{
			return Get-AccessToken -ClientId $ClientId -MailboxName $MailboxName -redirectUrl $redirectUrl -ClientSecret $ClientSecret -ResourceURL $ResourceURL -Beta:$beta.IsPresent -prompt $Prompt -CacheCredentials:$CacheCredentials.isPresent
		}
	}
}
