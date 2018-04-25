function Copy-EXROneDriveItem
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
		[String]
		$OneDriveFilePath,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$TargetFolderPath,
        
        [Parameter(Position = 4, Mandatory = $true)]
		[String]
		$TargetFileName
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
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $OneDriveFilePath
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		$RequestURL =  (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $JSONOutput.Id + "')/copy")
		$JSONPost = "{`r`n" 
		$JSONPost +=  "     `"parentReference`": {`r`n"
		if(![String]::IsNullOrEmpty($TargetFolderPath)){
			$TargetRequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $TargetFolderPath
			$TargetJSONOutput = Invoke-RestGet -RequestURL $TargetRequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			$JSONPost +=  "   `"driveId`": `"" + $TargetJSONOutput.parentReference.driveId + "`",`r`n"
        	$JSONPost +=  " `"id`": `"" + $TargetJSONOutput.id + "`"`r`n"
		}
		else{
			 $JSONPost +=  "   `"driveId`": `"" + $JSONOutput.parentReference.driveId + "`",`r`n"
        	 $JSONPost +=  " `"id`": `"" + $JSONOutput.parentReference.id + "`"`r`n"
		}
        $JSONPost += " },`r`n"
        $JSONPost += "`"name`": `"" + $TargetFileName + "`"`r`n}"
        $JSONOutput = Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $JSONPost
        
	}
}
