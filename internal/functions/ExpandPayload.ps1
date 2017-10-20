function ExpandPayload
{
	[CmdletBinding()]
	Param (
		$response
	)
	## Start Code Attribution
	## ExpandPayload function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://www.powershellgallery.com/profiles/chriswahl/
	## End Code Attribution
	[void][System.Reflection.Assembly]::LoadWithPartialName('System.Web.Extensions')
	return ParseItem -jsonItem ((New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer -Property @{
				MaxJsonLength  = [Int32]::MaxValue
			}).DeserializeObject($response))
}
