function Get-EXRBatchItems
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
		$Items,
		[Parameter(Position =3, Mandatory = $false)]
		[psobject]
		$SelectProperties,
		[Parameter(Position =4, Mandatory = $false)]
		[psobject]
		$PropList,		
		  
        [Parameter(Position = 5, Mandatory = $false)]
		[psobject]
		$URLString

	)
	Process
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
        $RequestURL = "https://graph.microsoft.com/v1.0/`$batch"
        $RequestContent = "{`r`n`"requests`": ["
        $itemCount = 1
        foreach($Item in $Items){
		   $ItemURI = $URLString + "('" + $Item.Id + "')"
		   $boolSelectProp = $false
		   if(![String]::IsNullOrEmpty($SelectProperties)){
				$ItemURI += "/?" +  $SelectProperties
				$boolSelectProp = $true
		   }
		   if($PropList -ne $null){
			   $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
			   if($boolSelectProp){
				   $ItemURI +=  "`&"
			   }else{
				    $ItemURI += "/?"
			   }
               $ItemURI +=  "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
           }
           if($itemCount -ne 1){$RequestContent += ",`r`n"}  
           $RequestContent += "{`r`n`"id`": `"" + $itemCount + "`",`r`n`"method`": `"GET`","
           $RequestContent += "`"url`": `"" + $ItemURI + "`"`r`n }"
           $itemCount++
		}
		$RequestContent += "`r`n]`r`n}"
		$JSONOutput = Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $RequestContent
		foreach($BatchItem in $JSONOutput.responses){
			Expand-ExtendedProperties -Item $BatchItem.Body
			Expand-MessageProperties -Item $BatchItem.Body
			Write-Output $BatchItem.Body
		}
	
	}
}
