function Set-EXRTracing
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[bool]
		$Tracing		
	)
	Begin
	{
		$Script:TraceRequest = $Tracing
	}
	
}
