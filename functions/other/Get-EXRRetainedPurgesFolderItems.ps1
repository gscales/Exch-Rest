function Get-EXRRetainedPurgesFolderItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName          
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
        return Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder recoverableitemspurges -PropList $Props
    }
}