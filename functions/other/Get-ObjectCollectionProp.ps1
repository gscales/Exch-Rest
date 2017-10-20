function Get-ObjectCollectionProp
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
		$CollectionProp = "" | Select-Object PropertyName, PropertyList, PropertyType, isArray
		$CollectionProp.PropertyType = "ObjectCollection"
		$CollectionProp.isArray = $false
		if ($Array.IsPresent) { $CollectionProp.isArray = $true }
		$CollectionProp.PropertyName = $Name
		if ($PropList -eq $null)
		{
			$CollectionProp.PropertyList = @()
		}
		else
		{
			$CollectionProp.PropertyList = $PropList
		}
		return, $CollectionProp
		
	}
}
