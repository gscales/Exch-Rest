function Invoke-RestPOST
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
		
		[Parameter(Position = 4, Mandatory = $true)]
		[PSCustomObject]
		$Content,

		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$BasicAuthentication,

		[Parameter(Position = 6, Mandatory = $false)]
		[System.Management.Automation.PSCredential]$Credentials,

		[Parameter(Position = 7, Mandatory = $false)]
		[string]
		$TimeZone
	


	)
	process
	{
		if($Script:TraceRequest){
			write-host $RequestURL
		}
		if($BasicAuthentication.IsPresent){
			$psString = $Credentials.UserName.ToString() + ":" + $Credentials.GetNetworkCredential().password.ToString()
			$psbyteArray = [System.Text.Encoding]::ASCII.($psString);
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
					if ($AccessToken.ADAL) {
                        if($AccessToken.ExpiresOn.DateTime -le [DateTime]::Now.ToUniversalTime().AddMinutes(10)) {
                            $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($AccessToken.Authority)
                            $token = ($Context.AcquireTokenSilentAsync($AccessToken.resource,$AccessToken.clientid)).Result
                            if($token.AccessToken -ne $AccessToken.AccessToken){
                                write-host "Refreshed Token ADAL"
                                if ([bool]($token.PSobject.Properties.name -match "AccessToken")) {
                                    #$AdalToken.access_token = 
                                    Add-Member -InputObject $token -NotePropertyName access_token -NotePropertyValue (Get-ProtectedToken -PlainToken $AccessToken.AccessToken) -Force
                                }
                                Add-Member -InputObject $token -NotePropertyName clientid -NotePropertyValue $AccessToken.clientid
                                Add-Member -InputObject $token -NotePropertyName ADAL -NotePropertyValue $True
                                Add-Member -InputObject $token -NotePropertyName redirectUrl -NotePropertyValue $AccessToken.redirectUrl
                                Add-Member -InputObject $token -NotePropertyName resource -NotePropertyValue $AccessToken.resource
                                Add-Member -InputObject $token -NotePropertyName mailbox -NotePropertyValue $AccessToken.mailbox
                                Add-Member -InputObject $token -NotePropertyName resourceCache -NotePropertyValue $AccessToken.resourceCache
                                if(![String]::IsNullOrEmpty($AccessToken.TenantId)){
                                    Add-Member -InputObject $token -NotePropertyName TenantId -NotePropertyValue $AccessToken.TenantId
                                }
                                Add-Member -InputObject $token -NotePropertyName Cached -NotePropertyValue $true
                                $HostDomain = (New-Object system.net.Mail.MailAddress($AccessToken.mailbox)).Host.ToLower()
                                if(!$Script:TokenCache[$AccessToken.resourceCache].ContainsKey($HostDomain)){			
                                    $Script:TokenCache[$AccessToken.resourceCache].Add($HostDomain,$token)
                                }
                                else{
                                    $Script:TokenCache[$AccessToken.resourceCache][$HostDomain] = $token
                                }
                                $AccessToken = $token
                            }
                        }
                    }else{
                        throw "App Token has expired"
                    }
				}
			}
			$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (ConvertFrom-SecureStringCustom -SecureToken $AccessToken.access_token));
		}
		$PostContent = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
		if([String]::IsNullOrEmpty($TimeZone)){
			$TimeZone = [TimeZoneInfo]::Local.Id
		}
		$HttpClient.DefaultRequestHeaders.Add("Prefer", ("outlook.timezone=`"" + $TimeZone + "`""))
		$ClientResult = $HttpClient.PostAsync([Uri]($RequestURL), $PostContent)
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
					if($JsonObject -ne $null){
						Add-Member -InputObject $JsonObject -NotePropertyName DateTimeRESTOperation -NotePropertyValue (Get-Date).ToString("s")
					}
					return $JsonObject
				}
				
			}
			
		}
		
	}
}
