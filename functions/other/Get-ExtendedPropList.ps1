function Get-ExtendedPropList
{
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[PSCustomObject]
		$PropertyList,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		$rtString = "";
		$PropName = "PropertyId"
		if ($AccessToken.resource -eq "https://graph.microsoft.com")
		{
			$PropName = "Id"
		}
		foreach ($Prop in $PropertyList)
		{
			if ($Prop.PropertyType -eq "Tagged")
			{
				if ($rtString -eq "")
				{
					$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"
				}
				else
				{
					$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"
				}
			}
			else
			{
				if ($Prop.Type -eq "String")
				{
					if ($rtString -eq "")
					{
						$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20" + $Prop.Id + "')"
					}
					else
					{
						$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20" + $Prop.Id + "')"
					}
				}
				else
				{
					if ($rtString -eq "")
					{
						$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20" + $Prop.Id + "')"
					}
					else
					{
						$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20" + $Prop.Id + "')"
					}
				}
			}
			
		}
		return $rtString
		
	}
}
