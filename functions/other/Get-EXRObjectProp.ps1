function Get-EXRObjectProp
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psObject]
		$PropList,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$Array
	)
	Begin
	{
		$ObjectProp = "" | Select-Object PropertyName, PropertyList, PropertyType, isArray
		$ObjectProp.PropertyType = "Object"
		$ObjectProp.isArray = $false
		if ($Array.IsPresent) { $ObjectProp.isArray = $true }
		$ObjectProp.PropertyName = $Name
		if ($PropList -eq $null)
		{
			$ObjectProp.PropertyList = @()
		}
		else
		{
			$ObjectProp.PropertyList = $PropList
		}
		return, $ObjectProp
		
	}
}
