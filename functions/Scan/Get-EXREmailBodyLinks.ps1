function Get-EXREmailBodyLinks {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$WellKnownFolder,
        [Parameter(Position = 2, Mandatory = $false)] [psobject]$Folder,
        [Parameter(Position = 3, Mandatory = $false)] [String]$FolderPath,
        [Parameter(Position = 4, Mandatory = $false)] [String]$MessageCount,
        [Parameter(Position = 5, Mandatory = $false)] [String]$InternetMessageId,
        [Parameter(Position=6, Mandatory=$false)] [switch]$SearchDumpster
    )
	
    process {
        $Props = @()
        $PR_BODY_HTML = Get-EXRTaggedProperty -DataType Binary -Id 0x1013
        $Props += $PR_BODY_HTML
        if([String]::IsNullOrEmpty($WellKnownFolder)){
			$WellKnownFolder = "AllItems"
			if($SearchDumpster.IsPresent){
				$WellKnownFolder = "RecoverableItemsDeletions"
			}
		} 
        if (![String]::IsNullOrEmpty($InternetMessageId)) {
            $Filter = "internetMessageId eq '" + $InternetMessageId + "'"
            $Item = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder $WellKnownFolder  -ReturnSize:$ReturnSize.IsPresent -SelectProperties $SelectProperties -Filter $Filter -Top $Top -OrderBy $OrderBy -TopOnly:$TopOnly.IsPresent -PropList $PropList   
            $BodyItem =  Get-exremail -ItemRESTURI  $Item.ItemRESTURI -PropList $Props
            Invoke-EXRParseEmailBodyLinks -Item $BodyItem -UseExtendedProperty
            Write-Output $BodyItem
        }
        else {
            Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder $WellKnownFolder -Folder $Folder -FolderPath $FolderPath -MessageCount $MessageCount -BatchReturnItems -SelectProperties Subject -PropList $Props | ForEach-Object {
                Invoke-EXRParseEmailBodyLinks -Item $_ -UseExtendedProperty
                Write-Output $_
            }
        }
       
       
    }
}