function ParseJsonObject
{
	[CmdletBinding()]
	Param (
		$jsonObj
	)
	## Start Code Attribution
	## ParseJsonObject function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://www.powershellgallery.com/profiles/chriswahl/
	## End Code Attribution
	$result = New-Object -TypeName PSCustomObject
	foreach ($key in $jsonObj.Keys)
	{
		$item = $jsonObj[$key]
		if ($item)
		{
			$parsedItem = ParseItem -jsonItem $item
		}
		else
		{
			$parsedItem = $null
		}
		$result | Add-Member -MemberType NoteProperty -Name $key -Value $parsedItem
	}
	return $result
}
