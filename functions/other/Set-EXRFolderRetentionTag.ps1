function Set-EXRFolderRetentionTag
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[String]
		$PolicyTagValue,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[Int32]
		$RetentionFlagsValue,
		
		[Parameter(Position = 5, Mandatory = $true)]
		[Int32]
		$RetentionPeriodValue
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
		$Folder = Get-EXRFolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($Folder -ne $null)
		{
			
			$retentionTagGUID = "{$($PolicyTagValue)}"
			$policyTagGUID = new-Object Guid($retentionTagGUID)
			$PolicyTagBase64 = [System.Convert]::ToBase64String($PolicyTagGUID.ToByteArray())
			$HttpClient = Get-HTTPClient -MailboxName $MailboxName
			$RequestURL = $Folder.FolderRestURI
			$FolderPostValue = "{`"SingleValueExtendedProperties`": [`r`n"
			$FolderPostValue += "`t{`"Id`":`"Binary 0x3019`",`"Value`":`"" + $PolicyTagBase64 + "`"},`r`n"
			$FolderPostValue += "`t{`"Id`":`"Integer 0x301D`",`"Value`":`"" + $RetentionFlagsValue + "`"},`r`n"
			$FolderPostValue += "`t{`"Id`":`"Integer 0x301A`",`"Value`":`"" + $RetentionPeriodValue + "`"}`r`n"
			$FolderPostValue += "]}"			
			return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $FolderPostValue
		}
	}
	
}
