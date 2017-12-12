function New-EXRDefaultAppRegistration
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$ClientId,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$RedirectURL
	)
	Process
	{
		$Config = "" | Select ClientId,RedirectURL
		$Config.ClientId = $ClientId
		$Config.RedirectURL = $RedirectURL
		$Directory = ($env:APPDATA + "\Exch-Rest")
	    if(!(Test-Path $Directory)){
			New-Item $Directory -type directory
		}
		Export-Clixml -InputObject $Config -Path ($env:APPDATA + "\Exch-Rest\appReg.cfg")
	}
}