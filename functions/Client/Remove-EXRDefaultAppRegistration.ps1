function Remove-EXRDefaultAppRegistration
{
	[CmdletBinding()]
	param (
	)
	Process
	{

		$File = ($env:APPDATA + "\Exch-Rest\appReg.cfg")
		if((Test-Path -Path $File)){
			Remove-item -Path $file -Confirm:$true
		}
	}
}