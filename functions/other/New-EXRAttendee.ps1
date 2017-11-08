function New-EXRAttendee
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$Address,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[ValidateSet("required", "optional", "resource")]
		[string]
		$Type
	)
	Begin
	{
		$Attendee = "" | Select-Object Name, Address, Type
		if ([String]::IsNullOrEmpty($Name))
		{
			$Attendee.Name = $Address
		}
		else
		{
			$Attendee.Name = $Name
		}
		$Attendee.Address = $Address
		$Attendee.Type = $Type
		return, $Attendee
	}
}
