function Get-MailAppProps
{
	[CmdletBinding()]
	Param (
		
	)
	
	#Holder for Mail Apps
	$cepPropdef = Get-NamedProperty -DataType "String" -Guid "00020329-0000-0000-C000-000000000046" -Id "cecp-propertyNames" -Value ($Guid + ";") -Type "String"
	$cepPropValue = Get-NamedProperty -DataType "String" -Guid "00020329-0000-0000-C000-000000000046" -Id "cecp-" + $Guid -Value $value -Type "String"
}
