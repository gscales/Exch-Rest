function Expand-ExtendedProperties
{
	[CmdletBinding()] 
    param (
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$Item
	)
	
 	process
	{
		if ($Item.singleValueExtendedProperties -ne $null)
		{
			foreach ($Prop in $Item.singleValueExtendedProperties)
			{
				Switch ($Prop.Id)
				{
                    "Binary 0x3019" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_POLICY_TAG" -NotePropertyValue ([System.GUID]([Convert]::FromBase64String($Prop.Value)))
                    }
                    "Binary 0x1013"{
                        Add-Member -InputObject $Item -NotePropertyName "PR_BODY_HTML" -NotePropertyValue ([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Prop.Value)))
                    }
                    "Binary 0xfff" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_ENTRYID" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-",""))
                    }
                    "Binary 0x301B" {
                        $fileTime = [BitConverter]::ToInt64([Convert]::FromBase64String($Prop.Value), 4);
                        $StartTime = [DateTime]::FromFileTime($fileTime)
                        Add-Member -InputObject $Item -NotePropertyName "PR_START_DATE_ETC" -NotePropertyValue $StartTime
                    }
                    "Binary 0x348A"{                            
                        Add-Member  -InputObject $Item -NotePropertyName "LastActiveParentEntryId" -NotePropertyValue ([System.BitConverter]::ToString([Convert]::FromBase64String($Prop.Value)).Replace("-",""))
                    }
                    "Integer 0x301D" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_RETENTION_FLAGS" -NotePropertyValue $Prop.Value
                    }
                    "Integer 0x301A" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_RETENTION_PERIOD" -NotePropertyValue $Prop.Value
                    }
                    "SystemTime 0x301C" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_RETENTION_DATE" -NotePropertyValue ([DateTime]::Parse($Prop.Value))
                    }
		             "String {403fc56b-cd30-47c5-86f8-ede9e35a022b} Name ComplianceTag" {
                        Add-Member -InputObject $Item -NotePropertyName "ComplianceTag" -NotePropertyValue $Prop.Value
                    }
                    "Integer {23239608-685D-4732-9C55-4C95CB4E8E33} Name InferenceClassificationResult" {
                        Add-Member -InputObject $Item -NotePropertyName "InferenceClassificationResult" -NotePropertyValue $Prop.Value
                    }
                    "Binary {e49d64da-9f3b-41ac-9684-c6e01f30cdfa} Name TeamChatFolderEntryId" {
                        Add-Member -InputObject $Item -NotePropertyName "TeamChatFolderEntryId" -NotePropertyValue $Prop.Value
                    }
                    "Integer 0xe08" {
                        Add-Member -InputObject $Item -NotePropertyName "Size" -NotePropertyValue $Prop.Value
                    }
                    "Long 0x66B3" {
                        Add-Member -InputObject $Item -NotePropertyName "FolderSize" -NotePropertyValue $Prop.Value
                    }
		            "String 0x7d" {
                        Add-Member -InputObject $Item -NotePropertyName "PR_TRANSPORT_MESSAGE_HEADERS" -NotePropertyValue $Prop.Value
                    }
                    "SystemTime 0xF02"{
                        Add-Member -InputObject $Item -NotePropertyName "PR_RENEWTIME" -NotePropertyValue ([DateTime]::Parse($Prop.Value))
                    }
                    "SystemTime 0xF01"{
                        Add-Member -InputObject $Item -NotePropertyName "PR_RENEWTIME2" -NotePropertyValue ([DateTime]::Parse($Prop.Value))
                    }
                    "String 0x66b5"{
                          Add-Member -InputObject $Item -NotePropertyName "PR_Folder_Path" -NotePropertyValue $Prop.Value.Replace("￾","\") -Force
                    }
                    "Short 0x3a4d"{
                          Add-Member -InputObject $Item -NotePropertyName "PR_Gender" -NotePropertyValue $Prop.Value -Force
                    }
                    "String 0x001a"{
                          Add-Member -InputObject $Item -NotePropertyName "PR_MESSAGE_CLASS" -NotePropertyValue $Prop.Value -Force
                    }
                    "Integer 0x6638"{
                          Add-Member -InputObject $Item -NotePropertyName "PR_FOLDER_CHILD_COUNT" -NotePropertyValue $Prop.Value -Force
                    }
                    "String {00062008-0000-0000-C000-000000000046} Name EntityExtraction/Sentiment1.0" {
                          Invoke-EXRProcessSentiment -Item $Item -JSONData $Prop.Value
                    }
                    "Integer {00062002-0000-0000-c000-000000000046} Id 0x8213" {
                        Add-Member -InputObject $Item -NotePropertyName "AppointmentDuration" -NotePropertyValue $Prop.Value -Force
                    }
                    default {Write-Host $Prop.Id}
                }
            }
        }
    }
}