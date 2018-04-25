function New-EXRExcelWorkSheet
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
		[psobject]
		$OneDriveFilePath,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
		$WorkSheetName 

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
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/drive/root:" + $OneDriveFilePath
		$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		$ItemURI =  (((Get-EndPoint -AccessToken $AccessToken -Segment "users") + "('$MailboxName')/drive") + "/items('" + $JSONOutput.Id + "')") + "/workbook/worksheets/add"
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$JSONPost = "{`r`n" 
        $JSONPost += "`"name`": `"" + $WorkSheetName + "`"`r`n}"
		$JSONOutput = Invoke-RestPOST -RequestURL $ItemURI -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $JSONPost
       	return $JSONOutput
	}
}
