function Get-EXRTaggedProperty
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$DataType,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Id,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Value
	)
	Begin
	{
		$Property = "" | Select-Object Id, DataType, PropertyType, Value
		$Property.Id = $Id
		$Property.DataType = $DataType
		$Property.PropertyType = "Tagged"
		if (![String]::IsNullOrEmpty($Value))
		{
			$Property.Value = $Value
		}
		return, $Property
	}
}
