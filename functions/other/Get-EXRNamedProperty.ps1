function Get-EXRNamedProperty
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$DataType,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Id,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Guid,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Type,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Value
	)
	Begin
	{
		$Property = "" | Select-Object Id, DataType, PropertyType, Type, Guid, Value
		$Property.Id = $Id
		$Property.DataType = $DataType
		$Property.PropertyType = "Named"
		$Property.Guid = $Guid
		if ($Type = "String")
		{
			$Property.Type = "String"
		}
		else
		{
			$Property.Type = "Id"
		}
		if (![String]::IsNullOrEmpty($Value))
		{
			$Property.Value = $Value
		}
		return, $Property
	}
}
