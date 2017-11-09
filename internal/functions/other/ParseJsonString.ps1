function ParseJsonString
{
	[CmdletBinding()]
	Param (
		$json
	)
	## Start Code Attribution
	## ParseJsonString function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://www.powershellgallery.com/profiles/chriswahl/
	## End Code Attribution
	$config = $javaScriptSerializer.DeserializeObject($json)
	return ParseJsonObject -jsonObj ($config)
}
