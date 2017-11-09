function Show-OAuthWindow
{
	[CmdletBinding()]
	param (
		[System.Uri]
		$Url
		
	)
	## Start Code Attribution
	## Show-AuthWindow function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://foxdeploy.com/2015/11/02/using-powershell-and-oauth/
	## https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/
	## End Code Attribution
	Add-Type -AssemblyName System.Web
	Add-Type -AssemblyName System.Windows.Forms
	
	$form = New-Object -TypeName System.Windows.Forms.Form -Property @{ Width = 440; Height = 640 }
	$web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{ Width = 420; Height = 600; Url = ($url) }
	$DocComp = {
		$Global:uri = $web.Url.AbsoluteUri
		if ($Global:Uri -match "error=[^&]*|code=[^&]*") { $form.Close() }
	}
	$web.ScriptErrorsSuppressed = $true
	$web.Add_DocumentCompleted($DocComp)
	$form.Controls.Add($web)
	$form.Add_Shown({ $form.Activate() })
	$form.ShowDialog() | Out-Null
	$queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
	$output = @{ }
	foreach ($key in $queryOutput.Keys)
	{
		$output["$key"] = $queryOutput[$key]
	}
	return $output
}
