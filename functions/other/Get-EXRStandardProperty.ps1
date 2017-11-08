function Get-EXRStandardProperty
{
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Id,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Value
	)
	Begin
	{
		$Property = "" | Select-Object Id, Value
		$Property.Id = $Id
		if (![String]::IsNullOrEmpty($Value))
		{
			$Property.Value = $Value
		}
		return, $Property
	}
}
