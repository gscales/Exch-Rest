
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
            "MessageSize"{
                $PidTagMessageSize = Get-EXRTaggedProperty -DataType "Integer" -Id "0x0E08"
                $PropList += $PidTagMessageSize 
            }
            "Sentiment"{
                $SentimentProp = Get-EXRNamedProperty -DataType "String" -Id "EntityExtraction/Sentiment1.0" -Type String -Guid "00062008-0000-0000-C000-000000000046"
                $PropList += $SentimentProp 
            }
                 
        }
        return $PropList
    }
}