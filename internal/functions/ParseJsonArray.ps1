function ParseJsonArray
{
	[CmdletBinding()]
	Param (
		$jsonArray
	)
	## Start Code Attribution
	## ParseJsonArray function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://www.powershellgallery.com/profiles/chriswahl/
	## End Code Attribution
	$result = @()
	$jsonArray | ForEach-Object -Process {
		$result += , (ParseItem -jsonItem $_)
	}
	return $result
}
