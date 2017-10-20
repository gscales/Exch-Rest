function Get-EXRItemRetentionTags(){
        $PR_POLICY_TAG = Get-TaggedProperty -DataType "Binary" -Id "0x3019"  
        $PR_RETENTION_FLAGS =  Get-TaggedProperty -DataType "Integer" -Id "0x301D" 
        $PR_RETENTION_PERIOD = Get-TaggedProperty -DataType "Integer" -Id "0x301A"   
        $PR_START_DATE_ETC  = Get-TaggedProperty -DataType "Binary" -Id "0x301B"  
        $PR_RETENTION_DATE   = Get-TaggedProperty -DataType "SystemTime" -Id "0x301C"  
        $ComplianceTag = Get-NamedProperty -DataType "String" -Id "ComplianceTag" -Type String -Guid '403FC56B-CD30-47C5-86F8-EDE9E35A022B'
        $Props = @()
        $Props +=$PR_POLICY_TAG
        $Props +=$PR_RETENTION_FLAGS
        $Props +=$PR_RETENTION_PERIOD
        $Props +=$PR_START_DATE_ETC
        $Props +=$PR_RETENTION_DATE 
	$Props +=$ComplianceTag
	return $Props
}