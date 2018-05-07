function Invoke-RestGet
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$RequestURL,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$MailboxName,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[System.Net.Http.HttpClient]
		$HttpClient,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[switch]
		$NoJSON,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[bool]
		$TrackStatus = $false,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[String]
		$ProcessMessage,

		[Parameter(Position = 7, Mandatory = $false)]
		[switch]
		$BasicAuthentication,

		[Parameter(Position = 8, Mandatory = $false)]
		[System.Management.Automation.PSCredential]$Credentials
	)
	process
	{
		if($BasicAuthentication.IsPresent){
			$psString = $Credentials.UserName.ToString() + ":" + $Credentials.GetNetworkCredential().password.ToString()
			$psbyteArray = [System.Text.Encoding]::ASCII.GetBytes($psString);
            $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Basic", [System.Convert]::ToBase64String($psbyteArray));
		}else{
			#Check for expired Token
			$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
			$expiry = $minTime.AddSeconds($AccessToken.expires_on)
			if ($expiry -le [DateTime]::Now.ToUniversalTime())
			{
				if ([bool]($AccessToken.PSobject.Properties.name -match "refresh_token"))
				{
					write-host "Refresh Token"
					$AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken	
				}
				else
				{
					throw "App Token has expired"
				}
				
			}
			if($Script:TraceRequest){
				write-host $RequestURL
			}
			$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $AccessToken.access_token));
		}
		$HttpClient.DefaultRequestHeaders.Add("Prefer", ("outlook.timezone=`"" + [TimeZoneInfo]::Local.Id + "`""))
		$ClientResult = $HttpClient.GetAsync($RequestURL)
		$exProgress = 0
		if ($TrackStatus)
		{
			While ($ClientResult.Status -eq [System.Threading.Tasks.TaskStatus]::Running -bor $ClientResult.Status -eq [System.Threading.Tasks.TaskStatus]::WaitingForActivation)
			{
				if ([String]::IsNullOrEmpty($ProcessMessage))
				{
					Write-Progress -Activity ("Executing Request " + $RequestURL) -PercentComplete $exProgress;
				}
				else
				{
					Write-Progress -Activity ($ProcessMessage) -PercentComplete $exProgress;
				}
				Start-Sleep -Milliseconds 200
				$exProgress += .2
			}
			Write-Progress -Activity "Executing Request" -Completed
		}
		if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
		{
			if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created)
			{
				write-Host ($ClientResult.Result)
			}
			if ($ClientResult.Result.Content -ne $null)
			{
				Write-Host ($ClientResult.Result.Content.ReadAsStringAsync().Result);
			}
		}
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Host ("Error making REST Get " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Host ("RequestURL : " + $RequestURL)
		}
		else
		{
			if ($NoJSON)
			{
				return $ClientResult.Result.Content
			}
			else
			{
				$JsonObject = ExpandPayload($ClientResult.Result.Content.ReadAsStringAsync().Result)
				#$JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
				if ([String]::IsNullOrEmpty($ClientResult))
				{
					write-host "No Value returned"
				}
				else
				{
					Add-Member -InputObject $JsonObject -NotePropertyName DateTimeRESTOperation -NotePropertyValue (Get-Date).ToString("s")
					return $JsonObject
				}
				
			}
			
		}
		
	}
}