function New-EXRReferanceAttachment
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
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
