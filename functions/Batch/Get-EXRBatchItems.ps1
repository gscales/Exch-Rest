function Get-EXRBatchItems {
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
        [Parameter(Position = 3, Mandatory = $false)]
        [psobject]
        $SelectProperties,
        [Parameter(Position = 4, Mandatory = $false)]
        [psobject]
        $PropList,		
		  
        [Parameter(Position = 5, Mandatory = $false)]
        [psobject]
        $URLString,

        [Parameter(Position = 6, Mandatory = $false)] 
        [switch]
        $ReturnAttachments,

        [Parameter(Position = 7, Mandatory = $false)] 
        [switch]
        $ProcessAntiSPAMHeaders,

        [Parameter(Position = 8, Mandatory = $false)] 
        [switch]
		$RestrictProps,
		
		[Parameter(Position = 9, Mandatory = $false)] 
        [switch]
        $ChildFolders

		
    )
    Process {
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        } 
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $RequestURL = "https://graph.microsoft.com/v1.0/`$batch"
        $RequestContent = "{`r`n`"requests`": ["
        $itemCount = 1
        foreach ($Item in $Items) {
			$ItemURI = $URLString + "('" + $Item.Id + "')"
			if($ChildFolders.IsPresent){
				$ItemURI +=  "/childfolders/?`$Top=1000"
			}
            $boolSelectProp = $false
            if ($RestrictProps.IsPresent) {
                if (![String]::IsNullOrEmpty($SelectProperties)) {
                    $ItemURI += "/?" + $SelectProperties
                    $boolSelectProp = $true
                }
            }
            if ($PropList -ne $null) {
                $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
                if ($boolSelectProp) {
                    $ItemURI += "`&"
                }
                else {
					if(!$ItemURI.Contains("/?")){
						$ItemURI += "/?"
					}                    
				}
				if($ChildFolders.IsPresent){
					$ItemURI += "`&"
				}
                $ItemURI += "`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            if ($itemCount -ne 1) {$RequestContent += ",`r`n"}  
            $RequestContent += "{`r`n`"id`": `"" + $itemCount + "`",`r`n`"method`": `"GET`","
            $RequestContent += "`"url`": `"" + $ItemURI + "`"`r`n }"
            $itemCount++
        }
        $RequestContent += "`r`n]`r`n}"
        $JSONOutput = Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $RequestContent
        foreach ($BatchItem in $JSONOutput.responses) {
            Expand-ExtendedProperties -Item $BatchItem.Body
            Expand-MessageProperties -Item $BatchItem.Body
            if ($ProcessAntiSPAMHeaders.IsPresent) {
                Invoke-EXRProcessAntiSPAMHeaders -Item $BatchItem.Body
            }
            if ($ReturnAttachments.IsPresent -band $Message.hasAttachments) {
                $AttachmentNames = @()
                $AttachmentDetails = @()
                Get-EXRAttachments -MailboxName $MailboxName -AccessToken $AccessToken -ItemURI $Message.ItemRESTURI | ForEach-Object {
                    $AttachmentNames += $_.name
                    $AttachmentDetails += $_    
                }
                add-Member -InputObject $BatchItem.Body -NotePropertyName AttachmentNames -NotePropertyValue $AttachmentNames
                add-Member -InputObject $BatchItem.Body -NotePropertyName AttachmentDetails -NotePropertyValue $AttachmentDetails
            }
            Write-Output $BatchItem.Body
        }
	
    }
}
