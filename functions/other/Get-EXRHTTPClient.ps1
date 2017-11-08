function Get-EXRHTTPClient
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName
	)
	Begin
	{
		Add-Type -AssemblyName System.Net.Http
		$handler = New-Object  System.Net.Http.HttpClientHandler
		$handler.CookieContainer = New-Object System.Net.CookieContainer
		$handler.AllowAutoRedirect = $true;
		$HttpClient = New-Object System.Net.Http.HttpClient($handler);
		#$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "");
		$Header = New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json")
		$HttpClient.DefaultRequestHeaders.Accept.Add($Header);
		$HttpClient.Timeout = New-Object System.TimeSpan(0, 0, 90);
		$HttpClient.DefaultRequestHeaders.TransferEncodingChunked = $false
		if (!$HttpClient.DefaultRequestHeaders.Contains("X-AnchorMailbox"))
		{
			$HttpClient.DefaultRequestHeaders.Add("X-AnchorMailbox", $MailboxName);
		}
		$Header = New-Object System.Net.Http.Headers.ProductInfoHeaderValue("RestClient", "1.1")
		$HttpClient.DefaultRequestHeaders.UserAgent.Add($Header);
		return $HttpClient
	}
}
