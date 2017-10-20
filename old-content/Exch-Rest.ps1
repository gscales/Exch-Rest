

function Get-AppSettings
{
	[CmdletBinding()]
	param (
		
	)
	Begin
	{
		$configObj = "" | Select-Object ResourceURL, ClientId, redirectUrl, ClientSecret, x5t, TenantId, ValidateForMinutes
		$configObj.ResourceURL = "outlook.office.com"
		$configObj.ClientId = "" # 1bdbfb41-f690-4f93-b0bb-002004bbca79
		$configObj.redirectUrl = "" # http://localhost:8000/authorize
		$configObj.TenantId = "" # 1c3a18bf-da31-4f6c-a404-2c06c9cf5ae4
		$configObj.ClientSecret = ""
		$configObj.x5t = "" # VS/H6cNa/3gc9FrSxGs9jOOZP3o=
		$configObj.ValidateForMinutes = 60
		return $configObj
	}
}

function Get-HTTPClient
{
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

function Convert-FromBase64StringWithNoPadding
{
	[CmdletBinding()]
	Param (
		[string]
		$Data
	)
	$data = $data.Replace('-', '+').Replace('_', '/')
	switch ($data.Length % 4)
	{
		0 { break }
		2 { $data += '==' }
		3 { $data += '=' }
		default { throw New-Object ArgumentException('data') }
	}
	return [System.Convert]::FromBase64String($data)
}

function Invoke-DecodeToken
{
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Token
	)
	## Start Code Attribution
	## Decode-Token function is based on work of the following Authors and should remain with the function if copied into other scripts
	## https://gallery.technet.microsoft.com/JWT-Token-Decode-637cf001
	## End Code Attribution
	Begin
	{
		$parts = $Token.Split('.');
		$headers = [System.Text.Encoding]::UTF8.GetString((Convert-FromBase64StringWithNoPadding $parts[0]))
		$claims = [System.Text.Encoding]::UTF8.GetString((Convert-FromBase64StringWithNoPadding $parts[1]))
		$signature = (Convert-FromBase64StringWithNoPadding $parts[2])
		
		$customObject = [PSCustomObject]@{
			headers  = ($headers | ConvertFrom-Json)
			claims   = ($claims | ConvertFrom-Json)
			signature = $signature
		}
		return $customObject
	}
}

function New-JWTToken
{
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$CertFileName,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$TenantId,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$ClientId,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[Int32]
		$ValidateForMinutes,
		
		[Parameter(Mandatory = $True)]
		[Security.SecureString]
		$password
	)
	Begin
	{
		
		$date1 = Get-Date -Date "01/01/1970"
		$date2 = (Get-Date).ToUniversalTime().AddMinutes($ValidateForMinutes)
		$date3 = (Get-Date).ToUniversalTime().AddMinutes(-5)
		$exp = [Math]::Round((New-TimeSpan -Start $date1 -End $date2).TotalSeconds, 0)
		$nbf = [Math]::Round((New-TimeSpan -Start $date1 -End $date3).TotalSeconds, 0)
		$exVal = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
		$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertFileName, $password, $exVal
		$x5t = [System.Convert]::ToBase64String($cert.GetCertHash())
		$jti = [System.Guid]::NewGuid().ToString()
		$Headerassertaion = "{"
		$Headerassertaion += "     `"alg`": `"RS256`","
		$Headerassertaion += "     `"x5t`": `"" + $x5t + "`""
		$Headerassertaion += "}"
		$PayLoadassertaion += "{"
		$PayLoadassertaion += "    `"aud`": `"https://login.windows.net/" + $TenantId + "/oauth2/token`","
		$PayLoadassertaion += "    `"exp`": $exp,"
		$PayLoadassertaion += "    `"iss`": `"" + $ClientId + "`","
		$PayLoadassertaion += "    `"jti`": `"" + $jti + "`","
		$PayLoadassertaion += "    `"nbf`": $nbf,"
		$PayLoadassertaion += "    `"sub`": `"" + $ClientId + "`""
		$PayLoadassertaion += "} "
		$encodedHeader = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($Headerassertaion)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$encodedPayLoadassertaion = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($PayLoadassertaion)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$JWTOutput = $encodedHeader + "." + $encodedPayLoadassertaion
		$SigBytes = [System.Text.UTF8Encoding]::UTF8.GetBytes($JWTOutput)
		$rsa = $cert.PrivateKey;
		$sha256 = [System.Security.Cryptography.SHA256]::Create()
		$hash = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($encodedHeader + '.' + $encodedPayLoadassertaion));
		$sigform = New-Object System.Security.Cryptography.RSAPKCS1SignatureFormatter($rsa);
		$sigform.SetHashAlgorithm("SHA256");
		$sig = [System.Convert]::ToBase64String($sigform.CreateSignature($hash)).Replace('=', '').Replace('+', '-').Replace('/', '_')
		$JWTOutput = $encodedHeader + '.' + $encodedPayLoadassertaion + '.' + $sig
		Write-Output ($JWTOutput)
		
	}
}

function Invoke-CreateSelfSignedCert
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$CertName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$CertFileName,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$KeyFileName
	)
	Begin
	{
		$Cert = New-SelfSignedCertificate -certstorelocation cert:\currentuser\my -dnsname $CertName -Provider 'Microsoft Enhanced RSA and AES Cryptographic Provider'
		$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
		$CertPath = "cert:\currentuser\my\" + $Cert.Thumbprint.ToString()
		Export-PfxCertificate -cert $CertPath -FilePath $CertFileName -Password $SecurePassword
		$bin = $cert.RawData
		$base64Value = [System.Convert]::ToBase64String($bin)
		$bin = $cert.GetCertHash()
		$base64Thumbprint = [System.Convert]::ToBase64String($bin)
		$keyid = [System.Guid]::NewGuid().ToString()
		$jsonObj = @{ customKeyIdentifier = $base64Thumbprint; keyId = $keyid; type = "AsymmetricX509Cert"; usage = "Verify"; value = $base64Value }
		$keyCredentials = ConvertTo-Json @($jsonObj) | Out-File $KeyFileName
		Remove-Item $CertPath
		Write-Host ("Key written to " + $KeyFileName)
		
	}
	
}

function Show-OAuthWindow
{
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

function Get-AccessToken
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$ClientId,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$redirectUrl,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$ClientSecret,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$ResourceURL,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[switch]
		$Beta,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[String]
		$Prompt
		
	)
	Begin
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient($MailboxName)
		$AppSetting = Get-AppSettings
		if ($ClientId -eq $null)
		{
			$ClientId = $AppSetting.ClientId
		}
		if ($ClientSecret -eq $null)
		{
			$ClientSecret = $AppSetting.ClientSecret
		}
		if ($redirectUrl -eq $null)
		{
			$redirectUrl = [System.Web.HttpUtility]::UrlEncode($AppSetting.redirectUrl)
		}
		else
		{
			$redirectUrl = [System.Web.HttpUtility]::UrlEncode($redirectUrl)
		}
		if ([String]::IsNullOrEmpty($ResourceURL))
		{
			$ResourceURL = $AppSetting.ResourceURL
		}
		if ([String]::IsNullOrEmpty($Prompt))
		{
			$Prompt = "refresh_session"
		}
		
		$Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=$Prompt"
		$code = $Phase1auth["code"]
		$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
		if (![String]::IsNullOrEmpty($ClientSecret))
		{
			$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_secret=$ClientSecret&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
		}
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
		$JsonObject = ConvertFrom-Json -InputObject $ClientReesult.Result.Content.ReadAsStringAsync().Result
		if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
		{
			$JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
		{
			$JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "id_token"))
		{
			$JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
		}
		Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
		Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
		if ($Beta.IsPresent)
		{
			Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
		}
		return $JsonObject
	}
}

function Get-AccessTokenUserAndPass
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$ClientId,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$ResourceURL,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[System.Management.Automation.PSCredential]
		$Credentials,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[switch]
		$Beta
	)
	Begin
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient($MailboxName)
		$AppSetting = Get-AppSettings
		if ($ClientId -eq $null)
		{
			$ClientId = $AppSetting.ClientId
		}
		
		if ([String]::IsNullOrEmpty($ResourceURL))
		{
			$ResourceURL = $AppSetting.ResourceURL
		}
		$UserName = $Credentials.UserName.ToString()
		$password = $Credentials.GetNetworkCredential().password.ToString()
		$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=password&username=$username&password=$password"
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
		$JsonObject = ConvertFrom-Json -InputObject $ClientReesult.Result.Content.ReadAsStringAsync().Result
		Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
		if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
		{
			$JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
		{
			$JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
		}
		if ($Beta.IsPresent)
		{
			Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "id_token"))
		{
			$JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
		}
		#Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
		return $JsonObject
	}
}

function Get-AppOnlyToken
{
	param (
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$CertFileName,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$TenantId,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$ClientId,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$redirectUrl,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[Int32]
		$ValidateForMinutes,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[string]
		$ResourceURL,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[switch]
		$Beta,
		
		[Parameter(Mandatory = $true)]
		[Security.SecureString]
		$password
		
	)
	Begin
	{
		$AppSetting = Get-AppSettings
		if ($TenantId -eq $null)
		{
			$AppSetting.TenantId
		}
		if ($ClientId -eq $null)
		{
			$ClientId = $AppSetting.ClientId
		}
		if ($redirectUrl -eq $null)
		{
			$redirectUrl = $AppSetting.redirectUrl
		}
		if ($ValidateForMinutes -eq 0)
		{
			$ValidateForMinutes = $AppSetting.ValidateForMinutes
		}
		if ([String]::IsNullOrEmpty($ResourceURL))
		{
			$ResourceURL = $AppSetting.ResourceURL
		}
		$JWTToken = New-JWTToken -CertFileName $CertFileName -password $password -TenantId $TenantId -ClientId $ClientId -ValidateForMinutes $ValidateForMinutes
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient(" ")
		$AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_assertion_type=urn%3Aietf%3Aparams%3Aoauth%3Aclient-assertion-type%3Ajwt-bearer&client_assertion=$JWTToken&grant_type=client_credentials&redirect_uri=$redirectUrl"
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/" + $TenantId + "/oauth2/token"), $content)
		$JsonObject = ConvertFrom-Json -InputObject $ClientReesult.Result.Content.ReadAsStringAsync().Result
		if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
		{
			$JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
		{
			$JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
		}
		if ([bool]($JsonObject.PSobject.Properties.name -match "id_token"))
		{
			$JsonObject.id_token = (Get-ProtectedToken -PlainToken $JsonObject.id_token)
		}
		Add-Member -InputObject $JsonObject -NotePropertyName tenantid -NotePropertyValue $TenantId
		Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
		Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
		if ($Beta.IsPresent)
		{
			Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
		}
		return $JsonObject
	}
}

function Invoke-RefreshAccessToken
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		Add-Type -AssemblyName System.Web
		$HttpClient = Get-HTTPClient($MailboxName)
		$ClientId = $AccessToken.clientid
		# $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AccessToken.redirectUrl)
		$redirectUrl = $AccessToken.redirectUrl
		$RefreshToken = (Get-TokenFromSecureString -SecureToken $AccessToken.refresh_token)
		$AuthorizationPostRequest = "client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"
		$content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
		$ClientResult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"), $content)
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			$JsonObject = ConvertFrom-Json -InputObject $ClientResult.Result.Content.ReadAsStringAsync().Result
			Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $AccessToken.clientid
			Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $AccessToken.redirectUrl
			if ([bool]($AccessToken.PSobject.Properties.name -match "Beta"))
			{
				Add-Member -InputObject $JsonObject -NotePropertyName Beta -NotePropertyValue True
			}
			if ([bool]($JsonObject.PSobject.Properties.name -match "refresh_token"))
			{
				$JsonObject.refresh_token = (Get-ProtectedToken -PlainToken $JsonObject.refresh_token)
			}
			if ([bool]($JsonObject.PSobject.Properties.name -match "access_token"))
			{
				$JsonObject.access_token = (Get-ProtectedToken -PlainToken $JsonObject.access_token)
			}
			return $JsonObject
		}
		
	}
}

function Get-TokenFromSecureString
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[System.Security.SecureString]
		$SecureToken
	)
	begin
	{
		$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
		$Token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
		return, $Token
	}
}

function Get-ProtectedToken
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$PlainToken
	)
	begin
	{
		$SecureString = New-Object System.Security.SecureString
		for ($i = 0; $i -lt $PlainToken.length; $i++)
		{
			$SecureString.AppendChar($PlainToken[$i])
		}
		$EncryptedToken = ConvertFrom-SecureString -SecureString $SecureString
		$SecureEncryptedToken = ConvertTo-SecureString -String $EncryptedToken
		return, $SecureEncryptedToken
	}
}

function ExpandPayload
{
	[CmdletBinding()]
	Param (
		$response
	)
	## Start Code Attribution
	## ExpandPayload function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://www.powershellgallery.com/profiles/chriswahl/
	## End Code Attribution
	[void][System.Reflection.Assembly]::LoadWithPartialName('System.Web.Extensions')
	return ParseItem -jsonItem ((New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer -Property @{
				MaxJsonLength  = [Int32]::MaxValue
			}).DeserializeObject($response))
}

function ParseItem
{
	[CmdletBinding()]
	Param (
		$JsonItem
	)
	
	if ($jsonItem.PSObject.TypeNames -match 'Array')
	{
		return ParseJsonArray -jsonArray ($jsonItem)
	}
	elseif ($jsonItem.PSObject.TypeNames -match 'Dictionary')
	{
		return ParseJsonObject -jsonObj ([HashTable]$jsonItem)
	}
	else
	{
		return $jsonItem
	}
}

function ParseJsonObject
{
	[CmdletBinding()]
	Param (
		$jsonObj
	)
	## Start Code Attribution
	## ParseJsonObject function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://www.powershellgallery.com/profiles/chriswahl/
	## End Code Attribution
	$result = New-Object -TypeName PSCustomObject
	foreach ($key in $jsonObj.Keys)
	{
		$item = $jsonObj[$key]
		if ($item)
		{
			$parsedItem = ParseItem -jsonItem $item
		}
		else
		{
			$parsedItem = $null
		}
		$result | Add-Member -MemberType NoteProperty -Name $key -Value $parsedItem
	}
	return $result
}

function ParseJsonArray
{
	[CmdletBinding()]
	Param (
		$jsonArray
	)
	## Start Code Attribution
	## ParseJsonArray function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://www.powershellgallery.com/profiles/chriswahl/
	## End Code Attribution
	$result = @()
	$jsonArray | ForEach-Object -Process {
		$result += , (ParseItem -jsonItem $_)
	}
	return $result
}

function ParseJsonString
{
	[CmdletBinding()]
	Param (
		$json
	)
	## Start Code Attribution
	## ParseJsonString function is the work of the following Authors and should remain with the function if copied into other scripts
	## https://www.powershellgallery.com/profiles/chriswahl/
	## End Code Attribution
	$config = $javaScriptSerializer.DeserializeObject($json)
	return ParseJsonObject -jsonObj ($config)
}

function Invoke-RestPOST
{
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
		
		[Parameter(Position = 3, Mandatory = $true)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[PSCustomObject]
		$Content
	)
	Begin
	{
		#Check for expired Token
		$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
		$expiry = $minTime.AddSeconds($AccessToken.expires_on)
		if ($expiry -le [DateTime]::Now.ToUniversalTime())
		{
			if ([bool]($AccessToken.PSobject.Properties.name -match "refresh_token"))
			{
				write-host "Refresh Token"
				$AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
				Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
			}
			else
			{
				throw "App Token has expired a new access token is required rerun get-apptoken"
			}
		}
		$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
		$PostContent = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
		$HttpClient.DefaultRequestHeaders.Add("Prefer", ("outlook.timezone=`"" + [TimeZoneInfo]::Local.Id + "`""))
		$ClientResult = $HttpClient.PostAsync([Uri]($RequestURL), $PostContent)
		if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
		{
			if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created)
			{
				write-Host ($ClientResult.Result)
			}
			if ($ClientResult.Result.Content -ne $null)
			{
				Write-Output ($ClientResult.Result.Content.ReadAsStringAsync());
			}
		}
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			$JsonObject = ExpandPayload($ClientResult.Result.Content.ReadAsStringAsync().Result)
			#$JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
			if ([String]::IsNullOrEmpty($JsonObject))
			{
				Write-Output $ClientResult.Result
			}
			else
			{
				return $JsonObject
			}
			
		}
		
	}
}

function Invoke-RestPatch
{
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
		
		[Parameter(Position = 3, Mandatory = $true)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[PSCustomObject]
		$Content
	)
	Begin
	{
		#Check for expired Token
		$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
		$expiry = $minTime.AddSeconds($AccessToken.expires_on)
		if ($expiry -le [DateTime]::Now.ToUniversalTime())
		{
			write-host "Refresh Token"
			$AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
			Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
		}
		$method = New-Object System.Net.Http.HttpMethod("PATCH")
		$HttpRequestMessage = New-Object System.Net.Http.HttpRequestMessage($method, [Uri]$RequestURL)
		$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
		$HttpRequestMessage.Content = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
		$ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
		if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
		{
			if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created)
			{
				write-Host ($ClientResult.Result)
			}
			if ($ClientResult.Result.Content -ne $null)
			{
				Write-Output ($ClientResult.Result.Content.ReadAsStringAsync());
			}
		}
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Output ("Error making REST Patch " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			# $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
			$JsonObject = ExpandPayload($ClientResult.Result.Content.ReadAsStringAsync().Result)
			if ([String]::IsNullOrEmpty($JsonObject))
			{
				Write-Output $ClientResult.Result
			}
			else
			{
				return $JsonObject
			}
			
		}
		
	}
}

function Invoke-RestPut
{
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
		
		[Parameter(Position = 3, Mandatory = $true)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[String]
		$ContentHeader,
		
		[Parameter(Position = 5, Mandatory = $true)]
		[PSCustomObject]
		$Content
		
	)
	Begin
	{
		#Check for expired Token
		$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
		$expiry = $minTime.AddSeconds($AccessToken.expires_on)
		if ($expiry -le [DateTime]::Now.ToUniversalTime())
		{
			write-host "Refresh Token"
			$AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
			Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
		}
		$method = New-Object System.Net.Http.HttpMethod("PUT")
		$HttpRequestMessage = New-Object System.Net.Http.HttpRequestMessage($method, [Uri]$RequestURL)
		$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
		if (![String]::IsNullorEmpty($ContentHeader))
		{
			$HttpRequestMessage.Content = New-Object System.Net.Http.ByteArrayContent -ArgumentList @( ,$Content)
		}
		else
		{
			$HttpRequestMessage.Content = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
		}
		$ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
		if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
		{
			if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created)
			{
				write-Host ($ClientResult.Result)
			}
			if ($ClientResult.Result.Content -ne $null)
			{
				Write-Output ($ClientResult.Result.Content.ReadAsStringAsync());
			}
		}
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Output ("Error making REST PUT " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			# $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
			$JsonObject = ExpandPayload($ClientResult.Result.Content.ReadAsStringAsync().Result)
			if ([String]::IsNullOrEmpty($JsonObject))
			{
				Write-Output $ClientResult.Result
			}
			else
			{
				return $JsonObject
			}
			
		}
		
	}
}

function Invoke-RestDELETE
{
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
		
		[Parameter(Position = 3, Mandatory = $true)]
		[psobject]
		$AccessToken
		
	)
	Begin
	{
		#Check for expired Token
		$minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
		$expiry = $minTime.AddSeconds($AccessToken.expires_on)
		if ($expiry -le [DateTime]::Now.ToUniversalTime())
		{
			write-host "Refresh Token"
			$AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
			Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
		}
		$method = New-Object System.Net.Http.HttpMethod("DELETE")
		$HttpRequestMessage = New-Object System.Net.Http.HttpRequestMessage($method, [Uri]$RequestURL)
		$HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
		$ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
		if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK)
		{
			if ($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::NoContent)
			{
				write-Host ($ClientResult.Result)
			}
			if ($ClientResult.Result.Content -ne $null)
			{
				Write-Output ($ClientResult.Result.Content.ReadAsStringAsync());
			}
		}
		if (!$ClientResult.Result.IsSuccessStatusCode)
		{
			Write-Output ("Error making REST Delete " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
			Write-Output $ClientResult.Result
			if ($ClientResult.Content -ne $null)
			{
				Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);
			}
		}
		else
		{
			$JsonObject = ConvertFrom-Json -InputObject $ClientResult.Result.Content.ReadAsStringAsync().Result
			if ([String]::IsNullOrEmpty($JsonObject))
			{
				Write-Output $ClientResult.Result
			}
			else
			{
				return $JsonObject
			}
			
		}
		
	}
}

function Get-MailboxSettings
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}

function Get-AutomaticRepliesSettings
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
			
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings/AutomaticRepliesSetting"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}

function Get-MailboxTimeZone
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings/TimeZone"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}

function Get-FolderFromPath
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	process
	{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/msgfolderroot/childfolders?"
		#  $RootFolder = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		#Split the Search path into an array  
		$tfTargetFolder = $RootFolder
		$fldArray = $FolderPath.Split("\")
		#Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++)
		{
			#Perform search based on the displayname of each folder level
			$FolderName = $fldArray[$lint];
			$RequestURL = $RequestURL += "`$filter=DisplayName eq '$FolderName'"
			$tfTargetFolder = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			if ($tfTargetFolder.Value.displayname -match $FolderName)
			{
				$folderId = $tfTargetFolder.value.Id.ToString()
				$RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$folderId')/childfolders?"
			}
			else
			{
				throw ("Folder Not found")
			}
		}
		if ($tfTargetFolder.Value -ne $null)
		{
			$folderId = $tfTargetFolder.Value.Id.ToString()
			Add-Member -InputObject $tfTargetFolder.Value -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
			return, $tfTargetFolder.Value
		}
		else
		{
			throw ("Folder Not found")
		}
	}
}

function Get-RootMailFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/msgfolderroot"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}

function Get-Inbox
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}

function Get-DefaultCalendarFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/calendar"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}

function Get-CalendarFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$FolderName
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/calendars?`$filter=name eq '" + $FolderName + "'"
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value)
			{
				Write-Output $Message
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
	}
}

function Get-AllCalendarFolders
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/calendars"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}

function Get-InboxItems
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,InferenceClassification`&`$Top=1000"
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value)
			{
				Write-Output $Message
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
	}
}

function Get-FocusedInboxItems
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,InferenceClassification`&`$Top=1000`&`$filter=InferenceClassification eq 'Focused'"
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value)
			{
				Write-Output $Message
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
	}
}

function Get-FolderItems
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[PSCustomObject]
		$Folder,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[switch]
		$ReturnSize,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[string]
		$SelectProperties,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[string]
		$Filter,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[string]
		$Top,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[string]
		$OrderBy,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[bool]
		$TopOnly
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		if (![String]::IsNullorEmpty($FolderPath))
		{
			$Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		}
		if (![String]::IsNullorEmpty($Filter))
		{
			$Filter = "`&`$filter=" + $Filter
		}
		if (![String]::IsNullorEmpty($Orderby))
		{
			$OrderBy = "`&`$OrderBy=" + $OrderBy
		}
		$TopValue = "1000"
		if (![String]::IsNullorEmpty($Top))
		{
			$TopValue = $Top
		}
		if ([String]::IsNullorEmpty($SelectProperties))
		{
			$SelectProperties = "`$select=ReceivedDateTime,Sender,Subject,IsRead"
		}
		else
		{
			$SelectProperties = "`$select=" + $SelectProperties
		}
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $Folder.FolderRestURI + "/messages/?" + $SelectProperties + "`&`$Top=" + $TopValue + $Filter + $OrderBy
			
			if ($ReturnSize.IsPresent)
			{
				$PropName = "PropertyId"
				if ($AccessToken.resource -eq "https://graph.microsoft.com")
				{
					$PropName = "Id"
				}
				$RequestURL = $Folder.FolderRestURI + "/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead`&`$Top=" + $TopValue + "`&`$expand=SingleValueExtendedProperties(`$filter=$PropName%20eq%20'Integer%200x0E08')" + $Filter + $OrderBy
			}
			write-host $RequestURL
			do
			{
				$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
				foreach ($Message in $JSONOutput.Value)
				{
					Add-Member -InputObject $Message -NotePropertyName ItemRESTURI -NotePropertyValue ($Folder.FolderRestURI + "/messages('" + $Message.Id + "')")
					Write-Output $Message
				}
				$RequestURL = $JSONOutput.'@odata.nextLink'
			}
			while (![String]::IsNullOrEmpty($RequestURL) -band (!$TopOnly))
		}
		
		
	}
}

function Move-Message
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[string]
		$TargetFolderPath
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		if ($TargetFolderPath -ne $null)
		{
			$Folder = Get-FolderFromPath -FolderPath $TargetFolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		}
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $ItemURI + "/move"
			$MoveItemPost = "{`"DestinationId`": `"" + $Folder.Id + "`"}"
			write-host $MoveItemPost
			return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $MoveItemPost
		}
	}
}

function Update-Message
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[String]
		$Subject,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[String]
		$Body,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[psobject]
		$Attachments,
		
		[Parameter(Position = 6, Mandatory = $false)]
		[psobject]
		$ToRecipients,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$StandardPropList,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[psobject]
		$ExPropList
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$RequestURL = $ItemURI
		$UpdateItemPatch = Get-MessageJSONFormat -Subject $Subject -Body $Body -Attachments $Attachments -ExPropList $ExPropList -StandardPropList $StandardPropList
		Write-host $UpdateItemPatch
		return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $UpdateItemPatch
	}
}

function Get-Attachments
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$ItemURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[switch]
		$MetaData,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[string]
		$SelectProperties
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		if ([String]::IsNullorEmpty($SelectProperties))
		{
			$SelectProperties = "`$select=Name,ContentType,Size,isInline,ContentType"
		}
		else
		{
			$SelectProperties = "`$select=" + $SelectProperties
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$RequestURL = $ItemURI + "/Attachments?" + $SelectProperties
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value)
			{
				Add-Member -InputObject $Message -NotePropertyName AttachmentRESTURI -NotePropertyValue ($ItemURI + "/Attachments('" + $Message.Id + "')")
				Write-Output $Message
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}

function Invoke-DownloadAttachment
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$AttachmentURI,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$AttachmentURI = $AttachmentURI + "?`$expand"
		$AttachmentObj = Invoke-RestGet -RequestURL $AttachmentURI -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		return $AttachmentObj
	}
}

function New-Folder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$ParentFolderPath,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$DisplayName
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$ParentFolder = Get-FolderFromPath -FolderPath $ParentFolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($ParentFolder -ne $null)
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $ParentFolder.FolderRestURI + "/childfolders"
			$NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"}"
			write-host $NewFolderPost
			return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
			
		}
		
		
	}
}

function New-ContactFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$DisplayName
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/ContactFolders"
		$NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"}"
		write-host $NewFolderPost
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
		
		
	}
}

function New-CalendarFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$DisplayName
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/calendars"
		$NewFolderPost = "{`"Name`": `"" + $DisplayName + "`"}"
		write-host $NewFolderPost
		return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
		
		
	}
}

function Rename-Folder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$NewDisplayName
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $Folder.FolderRestURI
			$RenameFolderPost = "{`"DisplayName`": `"" + $NewDisplayName + "`"}"
			return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $RenameFolderPost
			
		}
		
		
	}
}

function Update-FolderClass
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$FolderClass
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $Folder.FolderRestURI
			$UpdateFolderPost = "{`"SingleValueExtendedProperties`": [{`"PropertyId`":`"String 0x3613`",`"Value`":`"" + $FolderClass + "`"}]}"
			return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $UpdateFolderPost
			
		}
		
		
	}
}

function Update-Folder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$FolderPost
		
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($Folder -ne $null)
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $Folder.FolderRestURI
			$FolderPostValue = $FolderPost
			return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $FolderPostValue
			
		}
		
		
	}
}

function GetFolderRetentionTags
{
	[CmdletBinding()]
	Param (
		
	)
	
	#PR_POLICY_TAG 0x3019
	$PR_POLICY_TAG = Get-TaggedProperty -DataType "Binary" -Id "0x3019"
	#PR_RETENTION_FLAGS 0x301D   
	$PR_RETENTION_FLAGS = Get-TaggedProperty -DataType "Integer" -Id "0x301D"
	#PR_RETENTION_PERIOD 0x301A
	$PR_RETENTION_PERIOD = Get-TaggedProperty -DataType "Integer" -Id "0x301A"
}

function Set-FolderRetentionTag
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[String]
		$PolicyTagValue,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[Int32]
		$RetentionFlagsValue,
		
		[Parameter(Position = 5, Mandatory = $true)]
		[Int32]
		$RetentionPeriodValue
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($Folder -ne $null)
		{
			
			$retentionTagGUID = "{$($PolicyTagValue)}"
			$policyTagGUID = new-Object Guid($retentionTagGUID)
			$PolicyTagBase64 = [System.Convert]::ToBase64String($PolicyTagGUID.ToByteArray())
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $Folder.FolderRestURI
			$FolderPostValue = "{`"SingleValueExtendedProperties`": [`r`n"
			$FolderPostValue += "`t{`"PropertyId`":`"Binary 0x3019`",`"Value`":`"" + $PolicyTagBase64 + "`"},`r`n"
			$FolderPostValue += "`t{`"PropertyId`":`"Integer 0x301D`",`"Value`":`"" + $RetentionFlagsValue + "`"},`r`n"
			$FolderPostValue += "`t{`"PropertyId`":`"Integer 0x301A`",`"Value`":`"" + $RetentionPeriodValue + "`"}`r`n"
			$FolderPostValue += "]}"
			Write-Host $FolderPostValue
			return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $FolderPostValue
		}
	}
	
}

function Invoke-DeleteItem
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$ItemURI
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$confirmation = Read-Host "Are you Sure You Want To proceed with deleting the Item"
		if ($confirmation -eq 'y')
		{
			$HttpClient = Get-HTTPClient($MailboxName)
			$RequestURL = $ItemURI
			return Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		}
		else
		{
			Write-Host "skipped deletion"
		}
		
		
	}
}

function Invoke-DeleteFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FolderPath
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName
		if ($Folder -ne $null)
		{
			$confirmation = Read-Host "Are you Sure You Want To proceed with deleting Folder"
			if ($confirmation -eq 'y')
			{
				$HttpClient = Get-HTTPClient($MailboxName)
				$RequestURL = $Folder.FolderRestURI
				return Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			}
			else
			{
				Write-Host "skipped deletion"
			}
		}
		
		
	}
}

function Get-AllMailboxItems
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/AllItems/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,ParentFolderId`&`$Top=1000"
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value)
			{
				Write-Output $Message
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}

function Get-ExtendedPropList
{
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[PSCustomObject]
		$PropertyList,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		$rtString = "";
		$PropName = "PropertyId"
		if ($AccessToken.resource -eq "https://graph.microsoft.com")
		{
			$PropName = "Id"
		}
		foreach ($Prop in $PropertyList)
		{
			if ($Prop.PropertyType -eq "Tagged")
			{
				if ($rtString -eq "")
				{
					$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"
				}
				else
				{
					$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"
				}
			}
			else
			{
				if ($Prop.Type -eq "String")
				{
					if ($rtString -eq "")
					{
						$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20" + $Prop.Id + "')"
					}
					else
					{
						$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20" + $Prop.Id + "')"
					}
				}
				else
				{
					if ($rtString -eq "")
					{
						$rtString = "($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20" + $Prop.Id + "')"
					}
					else
					{
						$rtString += " or ($PropName%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20" + $Prop.Id + "')"
					}
				}
			}
			
		}
		return $rtString
		
	}
}

function Get-StandardProperty
{
	param (
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Id,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Value
	)
	Begin
	{
		$Property = "" | Select-Object Id, Value
		$Property.Id = $Id
		if (![String]::IsNullOrEmpty($Value))
		{
			$Property.Value = $Value
		}
		return, $Property
	}
}

function Get-TaggedProperty
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$DataType,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Id,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Value
	)
	Begin
	{
		$Property = "" | Select-Object Id, DataType, PropertyType, Value
		$Property.Id = $Id
		$Property.DataType = $DataType
		$Property.PropertyType = "Tagged"
		if (![String]::IsNullOrEmpty($Value))
		{
			$Property.Value = $Value
		}
		return, $Property
	}
}

function Get-NamedProperty
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[String]
		$DataType,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Id,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Guid,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[String]
		$Type,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[String]
		$Value
	)
	Begin
	{
		$Property = "" | Select-Object Id, DataType, PropertyType, Type, Guid, Value
		$Property.Id = $Id
		$Property.DataType = $DataType
		$Property.PropertyType = "Named"
		$Property.Guid = $Guid
		if ($Type = "String")
		{
			$Property.Type = "String"
		}
		else
		{
			$Property.Type = "Id"
		}
		if (![String]::IsNullOrEmpty($Value))
		{
			$Property.Value = $Value
		}
		return, $Property
	}
}

function Get-FolderClass
{
	[CmdletBinding()]
	Param (
		
	)
	
	$FolderClass = "" | Select-Object Id, DataType, PropertyType
	$FolderClass.Id = "0x3613"
	$FolderClass.DataType = "String"
	$FolderClass.PropertyType = "Tagged"
	return, $FolderClass
}

function Get-FolderPath
{
	[CmdletBinding()]
	Param (
		
	)
	$FolderPath = "" | Select-Object Id, DataType, PropertyType
	$FolderPath.Id = "0x66B5"
	$FolderPath.DataType = "String"
	$FolderPath.PropertyType = "Tagged"
	return, $FolderPath
}

function Get-AllMailFolders
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[PSCustomObject]
		$PropList
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders/msgfolderroot/childfolders?`$Top=1000"
		if ($PropList -ne $null)
		{
			$Props = Get-ExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
			$RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
			Write-Host $RequestURL
		}
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Folder in $JSONOutput.Value)
			{
				$Folder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ("\\" + $Folder.DisplayName)
				$folderId = $Folder.Id.ToString()
				Add-Member -InputObject $Folder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
				Write-Output $Folder
				if ($Folder.ChildFolderCount -gt 0)
				{
					if ($PropList -ne $null)
					{
						Get-AllChildFolders -Folder $Folder -AccessToken $AccessToken -PropList $PropList
					}
					else
					{
						Get-AllChildFolders -Folder $Folder -AccessToken $AccessToken
					}
				}
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}

function Get-AllChildFolders
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[PSCustomObject]
		$Folder,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[PSCustomObject]
		$PropList
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $Folder.FolderRestURI + "/childfolders/?`$Top=1000"
		if ($PropList -ne $null)
		{
			$Props = Get-ExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
			$RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
		}
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($ChildFolder in $JSONOutput.Value)
			{
				$ChildFolder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ($Folder.FolderPath + "\" + $ChildFolder.DisplayName)
				$folderId = $ChildFolder.Id.ToString()
				Add-Member -InputObject $ChildFolder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
				Write-Output $ChildFolder
				if ($ChildFolder.ChildFolderCount -gt 0)
				{
					if ($PropList -ne $null)
					{
						Get-AllChildFolders -Folder $ChildFolder -AccessToken $AccessToken -PropList $PropList
					}
					else
					{
						Get-AllChildFolders -Folder $ChildFolder -AccessToken $AccessToken
					}
				}
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}

function Get-AllCalendarFolders
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$FolderClass
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		if ($FolderClass.IsPresent)
		{
			$RequestURL = $EndPoint + "('$MailboxName')/Calendars/?`$Top=1000`&`$expand=SingleValueExtendedProperties(`$filter=PropertyId%20eq%20'String%200x66B5')"
			if ($AccessToken.resource -eq "https://graph.microsoft.com")
			{
				$RequestURL = $EndPoint + "('$MailboxName')/Calendars/?`$Top=1000`&`$expand=SingleValueExtendedProperties(`$filter=Id%20eq%20'String%200x66B5')"
			}
			
		}
		else
		{
			$RequestURL = $EndPoint + "('$MailboxName')/Calendars/?`$Top=1000"
		}
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Folder in $JSONOutput.Value)
			{
				Write-Output $Folder
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}

function Get-AllContactFolders
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/contactfolders/?`$Top=1000"
		Write-Host  $RequestURL
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Folder in $JSONOutput.Value)
			{
				Write-Output $Folder
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}

function Get-AllTaskfolders
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/taskfolders/?`$Top=1000"
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Folder in $JSONOutput.Value)
			{
				Write-Output $Folder
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
		
	}
}

function Get-ArchiveFolder
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings/ArchiveFolder"
		$JsonObject = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		$folderId = $JsonObject.value.ToString()
		$HttpClient = Get-HTTPClient($MailboxName)
		$RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$folderId')"
		return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
	}
}

function Get-MailboxSettingsReport
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[psobject]
		$Mailboxes,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$CertFileName,
		
		[Parameter(Mandatory = $True)]
		[Security.SecureString]
		$password
	)
	Begin
	{
		$rptCollection = @()
		$AccessToken = Get-AppOnlyToken -CertFileName $CertFileName -password $password
		$HttpClient = Get-HTTPClient($Mailboxes[0])
		foreach ($MailboxName in $Mailboxes)
		{
			$rptObj = "" | Select-Object MailboxName, Language, Locale, TimeZone, AutomaticReplyStatus
			$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
			$RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings"
			$Results = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			$rptObj.MailboxName = $MailboxName
			$rptObj.Language = $Results.Language.DisplayName
			$rptObj.Locale = $Results.Language.Locale
			$rptObj.TimeZone = $Results.TimeZone
			$rptObj.AutomaticReplyStatus = $Results.AutomaticRepliesSetting.Status
			$rptCollection += $rptObj
		}
		Write-Output  $rptCollection
		
	}
}

function Get-EndPoint
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[psObject]
		$AccessToken,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[psObject]
		$Segment,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[bool]
		$group
	)
	Begin
	{
		if ($group)
		{
			$Segment = "groups"
		}
		$EndPoint = "https://outlook.office.com/api/v2.0"
		switch ($AccessToken.resource)
		{
			"https://outlook.office.com" {
				if ($AccessToken.Beta)
				{
					$EndPoint = "https://outlook.office.com/api/beta/" + $Segment
				}
				else
				{
					$EndPoint = "https://outlook.office.com/api/v2.0/" + $Segment
				}
			}
			"https://graph.microsoft.com" {
				if ($AccessToken.Beta)
				{
					$EndPoint = "https://graph.microsoft.com/beta/" + $Segment
				}
				else
				{
					$EndPoint = "https://graph.microsoft.com/v1.0/" + $Segment
				}
			}
		}
		return, $EndPoint
		
	}
}

function Get-UserPhotoMetaData
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName + "/photo"
		$JsonObject = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Write-Output $JsonObject
	}
}

function Get-UserPhoto
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
		
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName + "/photo/`$value"
		$Result = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -NoJSON
		Write-Output $Result.ReadAsByteArrayAsync().Result
	}
}

function Get-MailboxUser
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName
		$JsonObject = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Write-Output $JsonObject
	}
}

function Get-CalendarGroups
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName + "/CalendarGroups"
		$JsonObject = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		Write-Output $JsonObject
	}
}

function Invoke-EnumCalendarGroups
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName + "/CalendarGroups"
		$JsonObject = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		foreach ($Group in $JsonObject.Value)
		{
			Write-Host ("GroupName : " + $Group.Name)
			$GroupId = $Group.Id.ToString()
			$RequestURL = $EndPoint + "/" + $MailboxName + "/CalendarGroups('$GroupId')/Calendars"
			$JsonObjectSub = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Calendar in $JsonObjectSub.Value)
			{
				Write-Host $Calendar.Name
			}
			$RequestURL = $EndPoint + "/" + $MailboxName + "/CalendarGroups('$GroupId')/MailFolders"
			$JsonObjectSub = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Calendar in $JsonObjectSub.Value)
			{
				Write-Host $Calendar.Name
			}
			
		}
		
		
	}
}

function Get-ObjectProp
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psObject]
		$PropList,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$Array
	)
	Begin
	{
		$ObjectProp = "" | Select-Object PropertyName, PropertyList, PropertyType, isArray
		$ObjectProp.PropertyType = "Object"
		$ObjectProp.isArray = $false
		if ($Array.IsPresent) { $ObjectProp.isArray = $true }
		$ObjectProp.PropertyName = $Name
		if ($PropList -eq $null)
		{
			$ObjectProp.PropertyList = @()
		}
		else
		{
			$ObjectProp.PropertyList = $PropList
		}
		return, $ObjectProp
		
	}
}

function Get-Recurrence
{
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$RecurrenceTimeZone,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[ValidateSet("daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly", " relativeYearly")]
		[string]
		$PatternType,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[Int]
		$PatternInterval,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[Int]
		$PatternMonth,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[Int]
		$PatternDayOfMonth,
		
		[Parameter(Position = 6, Mandatory = $true)]
		[ValidateSet("sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday")]
		[string]
		$PatternFirstDayOfWeek,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$PatternDaysOfWeek,
		
		[Parameter(Position = 8, Mandatory = $true)]
		[ValidateSet("first", "second", "third", "fourth", "last")]
		[string]
		$PatternIndex,
		
		[Parameter(Position = 9, Mandatory = $true)]
		[ValidateSet("noend", "enddate", "numbered")]
		[string]
		$RangeType,
		
		[Parameter(Position = 10, Mandatory = $true)]
		[datetime]
		$RangeStartDate,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[datetime]
		$RangeEndDate,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[Int]
		$RangeNumberOfOccurrences
	)
	Begin
	{
		$Recurrence = "" | Select-Object Pattern, Range, RecurrenceTimeZone
		$Pattern = "" | Select-Object Type, Interval, Month, DayOfMonth, DaysOfWeek, FirstDayOfWeek, Index
		$Range = "" | Select-Object  Type, StartDate, EndDate, NumberOfOccurrences
		if ([String]::IsNullOrEmpty($RecurrenceTimeZone))
		{
			$RecurrenceTimeZone = [TimeZoneInfo]::Local.Id
		}
		$Range.NumberOfOccurrences = 0
		$Pattern.Interval = 1
		$Pattern.Month = 0
		$Pattern.DayOfMonth = 0
		$Range.EndDate = "0001-01-01"
		$Recurrence.Pattern = $Pattern
		$Recurrence.Pattern.Type = $PatternType
		$Recurrence.Pattern.Interval = $PatternInterval
		if ($Recurrence.Pattern.Interval -eq 0)
		{
			$Recurrence.Pattern.Interval = 1
		}
		$Recurrence.Pattern.Month = $PatternMonth
		$Recurrence.Pattern.DayOfMonth = $PatternDayOfMonth
		$Recurrence.Pattern.DaysOfWeek = $PatternDaysOfWeek
		$Recurrence.Pattern.FirstDayOfWeek = $PatternFirstDayOfWeek
		$Recurrence.Pattern.Index = $PatternIndex
		$Recurrence.Range = $Range
		$Recurrence.Range.Type = $RangeType
		$Recurrence.Range.StartDate = $RangeStartDate.ToString("yyyy-MM-dd")
		if ($RangeEndDate -ne $null)
		{
			$Recurrence.Range.EndDate = $RangeEndDate.ToString("yyyy-MM-dd")
		}
		$Recurrence.Range.NumberOfOccurrences = $RangeNumberOfOccurrences
		$Recurrence.RecurrenceTimeZone = $RecurrenceTimeZone
		return, $Recurrence
	}
}

function Get-ObjectCollectionProp
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psObject]
		$PropList,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$Array
	)
	Begin
	{
		$CollectionProp = "" | Select-Object PropertyName, PropertyList, PropertyType, isArray
		$CollectionProp.PropertyType = "ObjectCollection"
		$CollectionProp.isArray = $false
		if ($Array.IsPresent) { $CollectionProp.isArray = $true }
		$CollectionProp.PropertyName = $Name
		if ($PropList -eq $null)
		{
			$CollectionProp.PropertyList = @()
		}
		else
		{
			$CollectionProp.PropertyList = $PropList
		}
		return, $CollectionProp
		
	}
}

function Get-ItemProp
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$Name,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[string]
		$Value,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[switch]
		$NoQuotes
	)
	Begin
	{
		$ItemProp = "" | Select-Object Name, Value, PropertyType, QuoteValue
		$ItemProp.PropertyType = "Single"
		$ItemProp.Name = $Name
		$ItemProp.Value = $Value
		if ($NoQuotes.IsPresent)
		{
			$ItemProp.QuoteValue = $false
		}
		else
		{
			$ItemProp.QuoteValue = $true
		}
		return, $ItemProp
		
	}
}

function Get-ModernGroups
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[string]
		$GroupName
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-AccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient($MailboxName)
		$RequestURL = Get-EndPoint -AccessToken $AccessToken -Segment "/groups?`$filter=groupTypes/any(c:c+eq+'Unified')"
		if (![String]::IsNullOrEmpty($GroupName))
		{
			$RequestURL = Get-EndPoint -AccessToken $AccessToken -Segment "/groups?`$filter=displayName eq '$GroupName'"
		}
		do
		{
			$JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Message in $JSONOutput.Value)
			{
				Write-Output $Message
			}
			$RequestURL = $JSONOutput.'@odata.nextLink'
		}
		while (![String]::IsNullOrEmpty($RequestURL))
		
	}
}

function Get-MailAppProps
{
	[CmdletBinding()]
	Param (
		
	)
	
	#Holder for Mail Apps
	$cepPropdef = Get-NamedProperty -DataType "String" -Guid "00020329-0000-0000-C000-000000000046" -Id "cecp-propertyNames" -Value ($Guid + ";") -Type "String"
	$cepPropValue = Get-NamedProperty -DataType "String" -Guid "00020329-0000-0000-C000-000000000046" -Id "cecp-" + $Guid -Value $value -Type "String"
}

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

function New-Attendee
{
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

function HexStringToByteArray
{
	[CmdletBinding()]
	Param (
		$HexString
	)
	
	$ByteArray = New-Object Byte[] ($HexString.Length/2);
	for ($i = 0; $i -lt $HexString.Length; $i += 2)
	{
		$ByteArray[$i/2] = [Convert]::ToByte($HexString.Substring($i, 2), 16)
	}
	Return @( ,$ByteArray)
	
}

