function Connect-EXRManagementAPI
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$UserName,
		
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
		$Outlook,

		[Parameter(Position = 9, Mandatory = $false)]
		[switch]
		$ShowMenu,

		[Parameter(Position = 10, Mandatory = $false)]
		[switch]
		$EnableTracing,

		[Parameter(Position = 11, Mandatory = $false)]
		[switch]
		$ManagementAPI


		
	)
	Begin
	{
		Connect-EXRMailbox -MailboxName $UserName -ClientId $ClientId -redirectUrl $redirectUrl -ClientSecret $ClientSecret -ResourceURL $ResourceURL -Prompt $Prompt -ShowMenu:$ShowMenu.IsPresent -ManagementAPI
	}
}
