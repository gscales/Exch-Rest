function Search-EXRContacts
{
	[CmdletBinding()]
	param (
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
		[Parameter(Position=3, Mandatory=$false)] [string]$FolderName,
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
		[Parameter(Position=208, Mandatory=$false)] [string]$Filter
		     
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
        if ([String]::IsNullOrEmpty($ContactsFolderName)) {
            $Contacts = Get-EXRDefaultContactsFolder -MailboxName $MailboxName -AccessToken $AccessToken
        }
        else {
            $Contacts = Get-EXRContactsFolder -MailboxName $MailboxName -AccessToken $AccessToken -FolderName $ContactsFolderName
            if ([String]::IsNullOrEmpty($Contacts)) {throw "Error Contacts folder not found check the folder name this is case sensitive"}
        }    
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users" 
        $RequestURL = $EndPoint + "('" + $MailboxName + "')/contactFolders('" + $Contacts.id + "')/contacts/?`$Top=1000"
        if(![String]::IsNullorEmpty($Filter)){
            $Filter = "`&`$filter=" + $Filter
        }
        if(![String]::IsNullorEmpty($Search)){
            $Search = "`&`$Search=`"" + $Search + "`""
        }
        if(![String]::IsNullorEmpty($Orderby)){
            $OrderBy = "`&`$OrderBy=" + $OrderBy
        }
        $TopValue = "1000"    
        if(![String]::IsNullorEmpty($Top)){
            $TopValue = $Top
        }      
        if([String]::IsNullorEmpty($SelectProperties)){
            $SelectProperties = "`$select=ReceivedDateTime,Sender,Subject,IsRead,hasAttachments"
        }
        else{
            $SelectProperties = "`$select=" + $SelectProperties
        }
		$HttpClient =  Get-HTTPClient -MailboxName $MailboxName
		if($ReturnSize.IsPresent){
			if($PropList -eq $null){
				$PropList = @()
				$PidTagMessageSize = Get-EXRTaggedProperty -DataType "Integer" -Id "0x0E08"  
				$PropList += $PidTagMessageSize
			}
		}
		if($PropList -ne $null){
			$Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
			$RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
		}
		$RequestURL += $Search + $Filter + $OrderBy
		do{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName 
			foreach ($Message in $JSONOutput.Value) {
				Add-Member -InputObject $Message -NotePropertyName ItemRESTURI -NotePropertyValue ($EndPoint + "('" + $MailboxName + "')"  + "/Contacts('" + $Message.Id + "')")
				Expand-ExtendedProperties -Item $Message
				Expand-MessageProperties -Item $Message
				if($ReturnAttachments.IsPresent -band $Message.hasAttachments){
					$AttachmentNames = @()
					$AttachmentDetails = @()
					Get-EXRAttachments -MailboxName $MailboxName -AccessToken $AccessToken -ItemURI $Message.ItemRESTURI | ForEach-Object{
						$AttachmentNames += $_.name
						$AttachmentDetails += $_    
					}
					add-Member -InputObject $Message -NotePropertyName AttachmentNames -NotePropertyValue $AttachmentNames
					add-Member -InputObject $Message -NotePropertyName AttachmentDetails -NotePropertyValue $AttachmentDetails
				}
				Write-Output $Message
			}           
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}while(![String]::IsNullOrEmpty($RequestURL) -band (!$TopOnly))     
        
		
		
		
	}
}
