function Get-EXRRetainedDeletedFolderItems{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [switch]$ReturnLastActiveParentFolderPath,
        [Parameter(Position=3, Mandatory=$false)] [switch]$Archive
    )
    Begin{
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
        $PR_POLICY_TAG = Get-EXRTaggedProperty -DataType "Binary" -Id "0x3019"  
        $PR_RETENTION_FLAGS =  Get-EXRTaggedProperty -DataType "Integer" -Id "0x301D" 
        $PR_RETENTION_PERIOD = Get-EXRTaggedProperty -DataType "Integer" -Id "0x301A"   
        $PR_START_DATE_ETC = Get-EXRTaggedProperty -DataType "Binary" -Id "0x301B"  
        $PR_RETENTION_DATE = Get-EXRTaggedProperty -DataType "SystemTime" -Id "0x301C"  
        $ComplianceTag = Get-EXRNamedProperty -DataType "String" -Id "ComplianceTag" -Type String -Guid '403FC56B-CD30-47C5-86F8-EDE9E35A022B'
        $Props = @()
        $Props +=$PR_POLICY_TAG
        $Props +=$PR_RETENTION_FLAGS
        $Props +=$PR_RETENTION_PERIOD
        $Props +=$PR_START_DATE_ETC
        $Props +=$PR_RETENTION_DATE 
        $Props +=$ComplianceTag
        $fldIndex = @{};
        if ($ReturnLastActiveParentFolderPath.IsPresent) {
            $Folders = Get-EXRAllMailFolders -MailboxName $MailboxName -AccessToken $AccessToken -ReturnEntryId
            foreach($folder in $Folders){
                $laFid = $folder.PR_ENTRYID.substring(44,44)
                Write-Verbose $laFid
                $fldIndex.Add($laFid,$folder);
            }
        } 
        if($Archive.IsPresent){
            $Items = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder archiverecoverableitemsDeletions -PropList $Props -ReturnLastActiveParentEntryId
        }else{
            $Items = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder recoverableitemsDeletions -PropList $Props -ReturnLastActiveParentEntryId
        }        
        if($ReturnLastActiveParentFolderPath.IsPresent){
            foreach($Item in $Items){
                if($Item.LastActiveParentEntryId){
                    if($fldIndex.ContainsKey($Item.LastActiveParentEntryId)){
                        $Item | Add-Member -Name "LastActiveParentFolderPath" -Value $fldIndex[$Item.LastActiveParentEntryId].FolderPath -MemberType NoteProperty
                        $Item | Add-Member -Name "LastActiveParentFolder" -Value $fldIndex[$Item.LastActiveParentEntryId] -MemberType NoteProperty
                    }
                }
            }
        }
        return  $Items
    }
}