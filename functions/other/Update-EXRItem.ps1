function Update-EXRItem {
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$details,

		[Parameter(Position = 4, Mandatory = $false)]
		[psobject]
		$PropList

	)
	Begin {
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
		if ($PropList -ne $null)
		{
			$details = "{"			
			$details += "`"SingleValueExtendedProperties`": [" + "`r`n"
			$propCount = 0
			$PropName = "PropertyId"
			if ($AccessToken.resource -eq "https://graph.microsoft.com")
			{
				$PropName = "Id"
			}
			foreach ($Property in $PropList)
			{
				if ($propCount -eq 0)
				{
					$details += "{" + "`r`n"
				}
				else
				{
					$NewMesdetailssage += ",{" + "`r`n"
				}
				if ($Property.PropertyType -eq "Tagged")
				{
					$details += "`"$PropName`":`"" + $Property.DataType + " " + $Property.Id + "`", " + "`r`n"
				}
				else
				{
					if ($Property.Type -eq "String")
					{
						$details += "`"$PropName`":`"" + $Property.DataType + " " + $Property.Guid + " Name " + $Property.Id + "`", " + "`r`n"
					}
					else
					{
						$details += "`"$PropName`":`"" + $Property.DataType + " " + $Property.Guid + " Id " + $Property.Id + "`", " + "`r`n"
					}
				}
				if($Property.Value -eq "null"){
					$details += "`"Value`":null" + "`r`n"
				}
				else{
					$details += "`"Value`":`"" + $Property.Value + "`"" + "`r`n"
				}				
				$details += " } " + "`r`n"
				$propCount++
			}
			$details += "]}" + "`r`n"
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$RequestURL = $ItemURI
		$results = Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $details
		return $results		
	}
}