function ParseItem
{
	[CmdletBinding()]
	Param (
		$JsonItem
	)
	
	if ($jsonItem.PSObject.TypeNames -match 'Array')
	{
		return ParseJsonArray -jsonArray ($jsonItem)
	}
	elseif ($jsonItem.PSObject.TypeNames -match 'Dictionary')
	{
		return ParseJsonObject -jsonObj ([HashTable]$jsonItem)
	}
	else
	{
		return $jsonItem
	}
}
