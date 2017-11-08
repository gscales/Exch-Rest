function Get-EXRMailAppProps
{
	[CmdletBinding()]
	Param (
		
	)
	
	#Holder for Mail Apps
	$cepPropdef = Get-EXRNamedProperty -DataType "String" -Guid "00020329-0000-0000-C000-000000000046" -Id "cecp-propertyNames" -Value ($Guid + ";") -Type "String"
	$cepPropValue = Get-EXRNamedProperty -DataType "String" -Guid "00020329-0000-0000-C000-000000000046" -Id "cecp-" + $Guid -Value $value -Type "String"
}
