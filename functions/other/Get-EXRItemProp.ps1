function Get-EXRItemProp
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$Value,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$NoQuotes
	)
	Begin
	{
		$ItemProp = "" | Select-Object Name, Value, PropertyType, QuoteValue
		$ItemProp.PropertyType = "Single"
		$ItemProp.Name = $Name
		$ItemProp.Value = $Value
		if ($NoQuotes.IsPresent)
		{
			$ItemProp.QuoteValue = $false
		}
		else
		{
			$ItemProp.QuoteValue = $true
		}
		return, $ItemProp
		
	}
}
