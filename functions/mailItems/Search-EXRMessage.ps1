function Search-EXRMessage
{
	[CmdletBinding()]
	param (
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [string]$WellKnownFolder,
        [Parameter(Position=4, Mandatory=$false)] [switch]$ReturnSize,
		[Parameter(Position=5, Mandatory=$false)] [string]$SelectProperties,
		[Parameter(Position=6, Mandatory=$false)] [string]$MessageId,
		[Parameter(Position=7, Mandatory=$false)] [string]$SubjectKQL,  
		[Parameter(Position=8, Mandatory=$false)] [string]$SubjectContains,  
		[Parameter(Position=8, Mandatory=$false)] [string]$SubjectStartsWith,
		[Parameter(Position=7, Mandatory=$false)] [string]$BodyKQL,  
		[Parameter(Position=8, Mandatory=$false)] [string]$BodyContains, 
		[Parameter(Position=9, Mandatory=$false)] [string]$KQL,
		[Parameter(Position=10, Mandatory=$false)] [DateTime]$ReceivedtimeFromKQL,
		[Parameter(Position=11, Mandatory=$false)] [DateTime]$ReceivedtimeToKQL,
		[Parameter(Position=12, Mandatory=$false)] [int]$First,
        [Parameter(Position=13, Mandatory=$false)] [PSCustomObject]$PropList     
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
		if(![String]::IsNullOrEmpty($MessageId)){
			$Filter = "internetMessageId eq '" + $MessageId + "'"
		}
		if(![String]::IsNullOrEmpty($SubjectContains)){
			$Filter = "contains(Subject,'" + $SubjectContains + "')"
		}
		if(![String]::IsNullOrEmpty($SubjectStartsWith)){
			$Filter = "startwith(Subject,'" + $SubjectStartsWith + "')"
		}
		if(![String]::IsNullOrEmpty($SubjectKQL)){
			$Search = "Subject: \`"" + $SubjectKQL + "\`""
		}
		if(![String]::IsNullOrEmpty($BodyContains)){
			$Filter = "contains(Body,'" + $BodyContains + "')"
		}
		if(![String]::IsNullOrEmpty($BodyKQL)){
			$Search = "Body:\`"" + $BodyKQL + "\`""
		}
		if(![String]::IsNullOrEmpty($KQL)){
			$Search = $KQL
		}
		if($ReceivedtimeFromKQL -ne $null -band $ReceivedtimeToKQL -ne $null){
			if([String]::IsNullOrEmpty($Search)){
				$Search = "Received:" + $ReceivedtimeFromKQL.ToString("yyyy-MM-dd") + ".." + $ReceivedtimeToKQL.ToString("yyyy-MM-dd")
			}
			else{
				$Search += "And Received:" + $ReceivedtimeFromKQL.ToString("yyyy-MM-dd") + ".." + $ReceivedtimeToKQL.ToString("yyyy-MM-dd")
			}
		}
		if($First -ne 0){
			$TopOnly = $true
			$Top = $First
		}
		else{
			$TopOnly = $false
		}
		Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder AllItems -ReturnSize:$ReturnSize.IsPresent -SelectProperties $SelectProperties -Search $Search -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly -PropList $PropList -ReturnFolderPath -ReturnStats
		
		
	}
}
