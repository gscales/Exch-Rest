function Get-EXRFolderPath
{
	[CmdletBinding()]
	Param (
		
	)
	$FolderPath = "" | Select-Object Id, DataType, PropertyType
	$FolderPath.Id = "0x66B5"
	$FolderPath.DataType = "String"
	$FolderPath.PropertyType = "Tagged"
	return, $FolderPath
}
