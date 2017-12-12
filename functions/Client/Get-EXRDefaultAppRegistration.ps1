function Get-EXRDefaultAppRegistration
{
	[CmdletBinding()]
	param (
	)
	Process
	{

		$File = ($env:APPDATA + "\Exch-Rest\appReg.cfg")
		if((Test-Path -Path $File)){
			return Import-Clixml -Path $file
		}
	}
}