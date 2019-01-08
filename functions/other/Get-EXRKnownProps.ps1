
function Get-EXRKnownProperty
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[String]
        $PropertyName,
        [Parameter(Position = 2, Mandatory = $false)]
		[Object[]]
        $PropList

	)
	
 	process
	{
        if($PropList -eq $null){
                 $PropList = @()
        }
        switch($PropertyName){
            "PR_ENTRYID" {
                $prEntryId = Get-EXRTaggedProperty -DataType "Binary" -Id "0x0FFF"  
                $PropList += $prEntryId
            }
            "PR_BODY_HTML" {
                $prEntryId = Get-EXRTaggedProperty -DataType "Binary" -Id "0x1013"  
                $PropList += $prEntryId
            }
            "PR_LAST_VERB_EXECUTED" {
                $prEntryId = Get-EXRTaggedProperty -DataType "Integer" -Id "0x1081"  
                $PropList += $prEntryId
            }
            "PR_LAST_VERB_EXECUTION_TIME" {
                $prEntryId = Get-EXRTaggedProperty -DataType "SystemTime" -Id "0x1082"  
                $PropList += $prEntryId
            }
            "MessageSize"{
                $PidTagMessageSize = Get-EXRTaggedProperty -DataType "Integer" -Id "0x0E08"
                $PropList += $PidTagMessageSize 
            }
            "FolderSize"{
                $FolderSizeProp = Get-EXRTaggedProperty -DataType "Long" -Id "0x66B3"
                $PropList += $FolderSizeProp 
            }
            "Sentiment"{
                $SentimentProp = Get-EXRNamedProperty -DataType "String" -Id "EntityExtraction/Sentiment1.0" -Type String -Guid "00062008-0000-0000-C000-000000000046"
                $PropList += $SentimentProp 
            }
            "LastActiveParentEntryId" {
                $LastActiveParentEntryId = Get-EXRTaggedProperty -DataType "Binary" -Id "0x348A"  
                $PropList += $LastActiveParentEntryId
            }
            "AppointmentDuration"{
                $AppointmentDuration = Get-EXRNamedProperty -DataType "Integer" -Id "0x8213" -Type Id -Guid "00062002-0000-0000-C000-000000000046"
                $PropList += $AppointmentDuration
            }
        }
        return $PropList
    }
}