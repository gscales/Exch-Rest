function Get-EXRFolderClass
{
	[CmdletBinding()]
	Param (
		
	)
	
	$FolderClass = "" | Select-Object Id, DataType, PropertyType
	$FolderClass.Id = "0x3613"
	$FolderClass.DataType = "String"
	$FolderClass.PropertyType = "Tagged"
	return, $FolderClass
}
