function GetFolderRetentionTags
{
	[CmdletBinding()]
	Param (
		
	)
	
	#PR_POLICY_TAG 0x3019
	$PR_POLICY_TAG = Get-TaggedProperty -DataType "Binary" -Id "0x3019"
	#PR_RETENTION_FLAGS 0x301D   
	$PR_RETENTION_FLAGS = Get-TaggedProperty -DataType "Integer" -Id "0x301D"
	#PR_RETENTION_PERIOD 0x301A
	$PR_RETENTION_PERIOD = Get-TaggedProperty -DataType "Integer" -Id "0x301A"
}
