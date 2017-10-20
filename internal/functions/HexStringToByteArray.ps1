function HexStringToByteArray
{
	[CmdletBinding()]
	Param (
		$HexString
	)
	
	$ByteArray = New-Object Byte[] ($HexString.Length/2);
	for ($i = 0; $i -lt $HexString.Length; $i += 2)
	{
		$ByteArray[$i/2] = [Convert]::ToByte($HexString.Substring($i, 2), 16)
	}
	Return @( ,$ByteArray)
	
}
