function New-EmailAddress
{
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$Address
	)
	Begin
	{
		$EmailAddress = "" | Select-Object Name, Address
		if ([String]::IsNullOrEmpty($Name))
		{
			$EmailAddress.Name = $Address
		}
		else
		{
			$EmailAddress.Name = $Name
		}
		$EmailAddress.Address = $Address
		return, $EmailAddress
	}
}
