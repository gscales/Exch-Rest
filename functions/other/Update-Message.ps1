function Update-Message
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$Subject,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[psobject]
		$ToRecipients,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$StandardPropList,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$ExPropList
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$RequestURL = $ItemURI
		$UpdateItemPatch = Get-MessageJSONFormat -Subject $Subject -Body $Body -Attachments $Attachments -ExPropList $ExPropList -StandardPropList $StandardPropList
		Write-host $UpdateItemPatch
		return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $UpdateItemPatch
	}
}
