function Invoke-UploadOneDriveItemToPath
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$OneDriveUploadFilePath,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$FilePath,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[Byte[]]
		$FileBytes
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $OneDriveUploadFilePath + ":/content"
		if ([String]::IsNullOrEmpty($FileBytes))
		{
			$Content = ([System.IO.File]::ReadAllBytes($filePath))
		}
		else
		{
			$Content = $FileBytes
		}
		$JSONOutput = Invoke-RestPut -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -content $Content -contentheader "application/octet-stream"
		return $JSONOutput
	}
}
