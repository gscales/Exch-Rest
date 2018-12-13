function Search-EXRMessage
{
	[CmdletBinding()]
	param (
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
		[Parameter(Position=2, Mandatory=$false)] [string]$WellKnownFolder,
		[Parameter(Position=3, Mandatory=$false)] [string]$FolderPath,
        [Parameter(Position=4, Mandatory=$false)] [switch]$ReturnSize,
		[Parameter(Position=5, Mandatory=$false)] [string]$SelectProperties,
		[Parameter(Position=6, Mandatory=$false)] [string]$MessageId,
		[Parameter(Position=7, Mandatory=$false)] [string]$Subject,  
		[Parameter(Position=7, Mandatory=$false)] [string]$SubjectKQL,  
		[Parameter(Position=8, Mandatory=$false)] [string]$SubjectContains,  
		[Parameter(Position=8, Mandatory=$false)] [string]$SubjectStartsWith,
		[Parameter(Position=7, Mandatory=$false)] [string]$BodyKQL,  
		[Parameter(Position=8, Mandatory=$false)] [string]$BodyContains, 
		[Parameter(Position=9, Mandatory=$false)] [string]$KQL,
		[Parameter(Position=9, Mandatory=$false)] [string]$From,
		[Parameter(Position=11, Mandatory=$false)] [string]$AttachmentKQL,
		[Parameter(Position=12, Mandatory=$false)] [DateTime]$ReceivedtimeFromKQL,
		[Parameter(Position=13, Mandatory=$false)] [DateTime]$ReceivedtimeToKQL,
		[Parameter(Position=14, Mandatory=$false)] [DateTime]$ReceivedtimeFrom,
		[Parameter(Position=15, Mandatory=$false)] [DateTime]$ReceivedtimeTo,
		[Parameter(Position=16, Mandatory=$false)] [int]$First,
		[Parameter(Position=17, Mandatory=$false)] [PSCustomObject]$PropList,
		[Parameter(Position=18, Mandatory=$false)] [switch]$ReturnStats,
		[Parameter(Position=19, Mandatory=$false)] [switch]$ReturnAttachments,
		[Parameter(Position=20, Mandatory=$false)] [string]$Filter,
		[Parameter(Position=21, Mandatory=$false)] [switch]$ReturnFolderPath,
        [Parameter(Position=24, Mandatory=$false)] [switch]$ReturnSentiment,
        [Parameter(Position=25, Mandatory=$false)] [switch]$ReturnEntryId,
        [Parameter(Position=26, Mandatory=$false)] [switch]$BatchReturnItems,
        [Parameter(Position=27, Mandatory=$false)] [switch]$ReturnInternetMessageHeaders,
        [Parameter(Position=28, Mandatory=$false)] [switch]$ProcessAntiSPAMHeaders
		     
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
		if(![String]::IsNullOrEmpty($KQL)){
			$Search = $KQL
		}
		if([String]::IsNullOrEmpty($MailboxName)){
			$MailboxName = $AccessToken.mailbox
		}  
		if(![String]::IsNullOrEmpty($MessageId)){
			$Filter = "internetMessageId eq '" + $MessageId + "'"
		}
		if(![String]::IsNullOrEmpty($Subject)){
			if([String]::IsNullOrEmpty($Filter)){
				$Filter = "Subject eq '" + $Subject + "'"
			}
			else{
				$Filter += " And Subject eq '" + $Subject + "'"
			}			
		}
		if(![String]::IsNullOrEmpty($SubjectContains)){
			if([String]::IsNullOrEmpty($Filter)){
				$Filter = "contains(Subject,'" + $SubjectContains + "')"
			}
			else{
				$Filter += " And contains(Subject,'" + $SubjectContains + "')"
			}			
		}
		if(![String]::IsNullOrEmpty($SubjectStartsWith)){
			if([String]::IsNullOrEmpty($Filter)){
				$Filter = "startwith(Subject,'" + $SubjectStartsWith + "')"
			}
			else{
				$Filter += " And startwith(Subject,'" + $SubjectStartsWith + "')"
			}				
		}
		if(![String]::IsNullOrEmpty($From)){
			if([String]::IsNullOrEmpty($Filter)){
				$Filter = "from/emailAddress/address eq '" + $From + "'"
			}
			else{
				$Filter += " And from/emailAddress/address eq '" + $From + "'"
			}			
		}
		if(![String]::IsNullOrEmpty($SubjectKQL)){
			$Search = "Subject: \`"" + $SubjectKQL + "\`""
		}
		if(![String]::IsNullOrEmpty($BodyContains)){
			$Filter = "contains(Body,'" + $BodyContains + "')"
		}
		if(![String]::IsNullOrEmpty($AttachmentKQL)){
			if([String]::IsNullOrEmpty($Search)){
				$Search = "attachment: '" + $AttachmentKQL + "'"
			}
			else{
				$Search += " And attachment: '" + $AttachmentKQL + "'"
			}
			
		}
		if(![String]::IsNullOrEmpty($ReceivedtimeFrom)){
			if([String]::IsNullOrEmpty($Filter)){
				$Filter = "receivedDateTime ge " + $ReceivedtimeFrom.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
			}
			else{
				$Filter += " And receivedDateTime ge " + $ReceivedtimeFrom.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
			}
		}		
		if(![String]::IsNullOrEmpty($ReceivedtimeTo)){
			if([String]::IsNullOrEmpty($Filter)){
				$Filter = "receivedDateTime le " + $ReceivedtimeTo.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
			}
			else{
				$Filter += " And receivedDateTime le " + $ReceivedtimeTo.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
			}
		}	
		if([String]::IsNullOrEmpty($WellKnownFolder)){
			$WellKnownFolder = "AllItems"
		}
		if(![String]::IsNullOrEmpty($BodyKQL)){
			if([String]::IsNullOrEmpty($Search)){
				$Search = "Body:\`"" + $BodyKQL + "\`""
			}
			else{
				$Search += " And Body:\`"" + $BodyKQL + "\`""
			}
			
		}

		if($ReceivedtimeFromKQL -ne $null -band $ReceivedtimeToKQL -ne $null){
			if([String]::IsNullOrEmpty($Search)){
				$Search = "Received:" + $ReceivedtimeFromKQL.ToString("yyyy-MM-dd") + ".." + $ReceivedtimeToKQL.ToString("yyyy-MM-dd")
			}
			else{
				$Search += " And Received:" + $ReceivedtimeFromKQL.ToString("yyyy-MM-dd") + ".." + $ReceivedtimeToKQL.ToString("yyyy-MM-dd")
			}
		}
		if($First -ne 0){
			$TopOnly = $true
			$Top = $First
		}
		else{
			$TopOnly = $false
		}
		if($ReturnStats.IsPresent){
			$DetailedStats = "" | Select TotalItems,TotalSize,TotalFolders,FolderStats
			$DetailedStats.TotalItems = 0
			$DetailedStats.TotalSize = 0 
			$DetailedStats.TotalFolders = 0
			$DetailedStats.FolderStats =  New-Object 'system.collections.generic.dictionary[[string],[Int32]]'
			if([String]::IsNullOrEmpty($FolderPath)){
				Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder $WellKnownFolder -ReturnSize:$true -SelectProperties $SelectProperties -Search $Search -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly -PropList $PropList -ReturnFolderPath -ReturnStats  -ReturnAttachments:$ReturnAttachments.IsPresent -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders | ForEach-Object{
					$DetailedStats.TotalItems++
					$DetailedStats.TotalSize += $_.Size 
				}
			}
			return $DetailedStats
		}
		else{
			if([String]::IsNullOrEmpty($FolderPath)){
				Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder $WellKnownFolder -ReturnSize:$ReturnSize.IsPresent -SelectProperties $SelectProperties -Search $Search -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly -PropList $PropList -ReturnFolderPath -ReturnStats -ReturnAttachments:$ReturnAttachments.IsPresent -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders
			}
			else{
				Get-EXRFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -FolderPath $FolderPath -ReturnSize:$ReturnSize.IsPresent -SelectProperties $SelectProperties -Search $Search -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly -PropList $PropList -ReturnAttachments:$ReturnAttachments.IsPresent -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders
			}
		}
		
		
		
	}
}
