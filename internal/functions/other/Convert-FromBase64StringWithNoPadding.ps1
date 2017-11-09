function Convert-FromBase64StringWithNoPadding
{
	[CmdletBinding()]
	Param (
		[string]
		$Data
	)
	$data = $data.Replace('-', '+').Replace('_', '/')
	switch ($data.Length % 4)
	{
		0 { break }
		2 { $data += '==' }
		3 { $data += '=' }
		default { throw New-Object ArgumentException('data') }
	}
	return [System.Convert]::FromBase64String($data)
}
