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
		[Parameter(Position=7, Mandatory=$false)] [string]$Subject,  
		[Parameter(Position=9, Mandatory=$false)] [int]$First,
        [Parameter(Position=10, Mandatory=$false)] [PSCustomObject]$PropList     
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
		if(![String]::IsNullOrEmpty($Subject)){
			$Search = "`"Subject:'" + $Subject + "'`""
		}
		if($First -ne 0){
			$TopOnly = $true
			$Top = $First
		}
		else{
			$TopOnly = $false
		}
		Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder AllItems -ReturnSize:$ReturnSize.IsPresent -SelectProperties $SelectProperties -Search $Search -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly -PropList $PropList -ReturnFolderPath
		
		
	}
}
