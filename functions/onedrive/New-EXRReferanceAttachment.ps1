function New-EXRReferanceAttachment
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$SourceUrl,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$ProviderType,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[String]
		$Permission,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$IsFolder
		
	)
	Begin
	{
		$ReferanceAttachment = "" | Select-Object Name, SourceUrl, ProviderType, Permission, IsFolder
		$ReferanceAttachment.IsFolder = "False"
		$ReferanceAttachment.ProviderType = "oneDriveBusiness"
		$ReferanceAttachment.Permission = $Permission
		$ReferanceAttachment.SourceUrl = $SourceUrl
		$ReferanceAttachment.Name = $Name
		if (![String]::IsNullOrEmpty($ProviderType))
		{
			$ReferanceAttachment.ProviderType = $ProviderType
		}
		if (![String]::IsNullOrEmpty($IsFolder))
		{
			$ReferanceAttachment.IsFolder = $IsFolder
		}
		return $ReferanceAttachment
	}
}
