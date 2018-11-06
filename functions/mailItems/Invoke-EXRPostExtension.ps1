function Invoke-EXRPostExtension
{
	[CmdletBinding()]
	param (

	    [parameter(ValueFromPipeline=$True)]
		[psobject[]]$Item,
		
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,

		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$ItemURI,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
        $AccessToken,
        
        [Parameter(Position = 4, Mandatory = $false)]
		[String]
        $extensionName,

        [Parameter(Position = 5, Mandatory = $false)]
		[psobject]
        $Values
        

	)
	Process
	{
		Write-Host $Item
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

        $RequestURL =  $ItemURI + "/extensions"        
        $Update = "{"
        $Update +=  "`"@odata.type`":`"microsoft.graph.openTypeExtension`","
        $Update +=  "`"extensionName`":`"$extensionName`","
        foreach($value in $Values){
            $Update +=  $value
        }        
        $Update += "}"

        
		return Invoke-RestPost -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $Update
	}
}
