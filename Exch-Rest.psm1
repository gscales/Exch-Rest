function Get-AppSettings(){
        param( 
        
        )  
 	Begin
		 {
            $configObj = "" |  select ResourceURL,ClientId,redirectUrl,ClientSecret,x5t,TenantId,ValidateForMinutes
            $configObj.ResourceURL = "outlook.office.com"
            $configObj.ClientId = "" # eg 1bdbfb41-f690-4f93-b0bb-002004bbca79
            $configObj.redirectUrl = "" # http://localhost:8000/authorize
            $configObj.TenantId = "" # eg 1c3a18bf-da31-4f6c-a404-2c06c9cf5ae4
            $configObj.ClientSecret = ""
            $configObj.x5t = "" # eg VS/H6cNa/3gc9FrSxGs9jOOZP3o=
            $configObj.ValidateForMinutes = 60
            return $configObj            
         }    
}

function Get-HTTPClient{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName
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
            if (!$HttpClient.DefaultRequestHeaders.Contains("X-AnchorMailbox")){
                $HttpClient.DefaultRequestHeaders.Add("X-AnchorMailbox", $MailboxName);
            }
            $Header = New-Object System.Net.Http.Headers.ProductInfoHeaderValue("RestClient", "1.1")
            $HttpClient.DefaultRequestHeaders.UserAgent.Add($Header);
            return $HttpClient
         }
}

function Convert-FromBase64StringWithNoPadding([string]$data)
{
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

function Decode-Token { 
        param( 
        [Parameter(Position=1, Mandatory=$true)] [String]$Token
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
            headers = ($headers | ConvertFrom-Json)
            claims = ($claims | ConvertFrom-Json)
            signature = $signature
        }
        return $customObject
    }
}

function New-JWTToken{
        param( 
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Position=2, Mandatory=$true)] [string]$TenantId,
        [Parameter(Position=3, Mandatory=$true)] [string]$ClientId,
        [Parameter(Position=4, Mandatory=$true)] [string]$x5t,
        [Parameter(Position=4, Mandatory=$true)] [Int32]$ValidateForMinutes,
        [Parameter(Mandatory=$True)][Security.SecureString]$password        
    )  
 	Begin
		 {
           
            $date1 = Get-Date -Date "01/01/1970"
            $date2 = (Get-Date).ToUniversalTime().AddMinutes($ValidateForMinutes)           
            $date3 = (Get-Date).ToUniversalTime().AddMinutes(-5)      
            $exp = [Math]::Round((New-TimeSpan -Start $date1 -End $date2).TotalSeconds,0) 
            $nbf = [Math]::Round((New-TimeSpan -Start $date1 -End $date3).TotalSeconds,0) 
            $exVal = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $CertFileName,$password,$exVal
            $jti = [System.Guid]::NewGuid().ToString()
            $Headerassertaion =  "{" 
            $Headerassertaion += "     `"alg`": `"RS256`"," 
            $Headerassertaion += "     `"x5t`": `""+ $x5t + "`""
            $Headerassertaion += "}"
            $PayLoadassertaion += "{"
            $PayLoadassertaion += "    `"aud`": `"https://login.windows.net/" + $TenantId +"/oauth2/token`"," 
            $PayLoadassertaion += "    `"exp`": $exp,"            
            $PayLoadassertaion += "    `"iss`": `""+ $ClientId + "`"," 
            $PayLoadassertaion += "    `"jti`": `"" + $jti + "`","
            $PayLoadassertaion += "    `"nbf`": $nbf,"       
            $PayLoadassertaion += "    `"sub`": `"" + $ClientId + "`""              
            $PayLoadassertaion += "} " 
            $encodedHeader = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($Headerassertaion)).Replace('=','').Replace('+', '-').Replace('/', '_')
            $encodedPayLoadassertaion = [System.Convert]::ToBase64String([System.Text.UTF8Encoding]::UTF8.GetBytes($PayLoadassertaion)).Replace('=','').Replace('+', '-').Replace('/', '_')
            $JWTOutput = $encodedHeader + "." + $encodedPayLoadassertaion
            $SigBytes = [System.Text.UTF8Encoding]::UTF8.GetBytes($JWTOutput)            
            $rsa = $cert.PrivateKey;
            $sha256 = [System.Security.Cryptography.SHA256]::Create()
            $hash = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($encodedHeader + '.' + $encodedPayLoadassertaion));
            $sigform = New-Object System.Security.Cryptography.RSAPKCS1SignatureFormatter($rsa);
            $sigform.SetHashAlgorithm("SHA256");
            $sig = [System.Convert]::ToBase64String($sigform.CreateSignature($hash)).Replace('=','').Replace('+', '-').Replace('/', '_')
            $JWTOutput = $encodedHeader + '.' + $encodedPayLoadassertaion + '.' + $sig
            Write-Output ($JWTOutput)

         }
}

function Invoke-CreateSelfSignedCert{
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$CertName,
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Position=2, Mandatory=$true)] [string]$KeyFileName
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
             $jsonObj = @{customKeyIdentifier=$base64Thumbprint;keyId=$keyid;type="AsymmetricX509Cert";usage="Verify";value=$base64Value}
             $keyCredentials=ConvertTo-Json @($jsonObj) | Out-File $KeyFileName
             Remove-Item $CertPath
             Write-Host ("Key written to " + $KeyFileName)
             
         }
    
}

Function Show-OAuthWindow
{
    param(
        [System.Uri]$Url
    )
    ## Start Code Attribution
    ## Show-AuthWindow function is the work of the following Authors and should remain with the function if copied into other scripts
    ## https://foxdeploy.com/2015/11/02/using-powershell-and-oauth/
    ## https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/
    ## End Code Attribution
    Add-Type -AssemblyName System.Web
    Add-Type -AssemblyName System.Windows.Forms
 
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url ) }
    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri
        if ($Global:Uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null
    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }
    return $output 
}

function Get-AccessToken{ 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [string]$ClientId,
        [Parameter(Position=2, Mandatory=$false)] [string]$redirectUrl,
        [Parameter(Position=3, Mandatory=$false)] [string]$ClientSecret
    )  
 	Begin
		 {
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient($MailboxName)
            $AppSetting = Get-AppSettings 
            $ResourceURL = $AppSetting.ResourceURL
            if($ClientId -eq $null){
                 $ClientId = $AppSetting.ClientId
            }
            if($ClientSecret -eq $null){
                 $ClientSecret =  $AppSetting.ClientSecret
            }           
            if($redirectUrl -eq $null){
                $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AppSetting.redirectUrl)
            }
            else{
                $redirectUrl = [System.Web.HttpUtility]::UrlEncode($redirectUrl)
            }
            $Phase1auth = Show-OAuthWindow -Url "https://login.microsoftonline.com/common/oauth2/authorize?resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&response_type=code&redirect_uri=$redirectUrl&prompt=login"
            $code = $Phase1auth["code"]
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            if(![String]::IsNullOrEmpty($ClientSecret)){
                $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_secret=$ClientSecret&grant_type=authorization_code&code=$code&redirect_uri=$redirectUrl"
            }
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"),$content)
            $JsonObject = ConvertFrom-Json -InputObject  $ClientReesult.Result.Content.ReadAsStringAsync().Result
            return $JsonObject
         }
}

function Get-AppOnlyToken{ 
    param( 
       
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Position=2, Mandatory=$false)] [string]$TenantId,
        [Parameter(Position=3, Mandatory=$false)] [string]$ClientId,
        [Parameter(Position=4, Mandatory=$false)] [string]$redirectUrl,     
        [Parameter(Position=5, Mandatory=$false)] [string]$x5t,
        [Parameter(Position=6, Mandatory=$false)] [Int32]$ValidateForMinutes,
        [Parameter(Mandatory=$true)] [Security.SecureString]$password
        
    )  
 	Begin
		 {
             $AppSetting = Get-AppSettings 
            if($TenantId -eq $null){
                $AppSetting.TenantId
            }
            if($ClientId -eq $null){
                 $ClientId = $AppSetting.ClientId
            }
            if($x5t -eq $null){
                 $x5t = $AppSetting.x5t
            }
            if($redirectUrl -eq $null){
                $redirectUrl = $AppSetting.redirectUrl
            }
            if($ValidateForMinutes -eq 0){
                $ValidateForMinutes = $AppSetting.ValidateForMinutes               
            }
            $JWTToken = New-JWTToken -CertFileName $CertFileName -password $password -TenantId $TenantId -ClientId $ClientId -x5t $x5t -ValidateForMinutes $ValidateForMinutes
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient(" ")          
            $ResourceURL = $AppSetting.ResourceURL
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_assertion_type=urn%3Aietf%3Aparams%3Aoauth%3Aclient-assertion-type%3Ajwt-bearer&client_assertion=$JWTToken&grant_type=client_credentials&redirect_uri=$redirectUrl"
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/" + $TenantId + "/oauth2/token"),$content)
            $JsonObject = ConvertFrom-Json -InputObject  $ClientReesult.Result.Content.ReadAsStringAsync().Result
            return $JsonObject
         }
}



function Refresh-AccessToken{ 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$RefreshToken
    )  
 	Begin
		 {
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient($MailboxName)
            $AppSetting = Get-AppSettings 
            $ResourceURL = $AppSetting.ResourceURL
            $ClientId = $AppSetting.ClientId
            $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AppSetting.redirectUrl)
            $AuthorizationPostRequest = "client_id=$ClientId&refresh_token=$RefreshToken&grant_type=refresh_token&redirect_uri=$redirectUrl"
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientResult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"),$content)             
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {                    
                     Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
               return $JsonObject
             }

         }
}

function Invoke-RestGet
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [PSCustomObject]$AccessToken
    )  
 	Begin
		 {
             #Check for expired Token
             $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0,[System.DateTimeKind]::Utc);
             $expiry =  $minTime.AddSeconds($AccessToken.expires_on)
             if($expiry -le [DateTime]::Now.ToUniversalTime()){
                if([bool]($AccessToken.PSobject.Properties.name -match "refresh_token")){
                    write-host "Refresh Token"
                    $AccessToken = Refresh-AccessToken -MailboxName $MailboxName -RefreshToken $AccessToken.refresh_token               
                    Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
                }
                else{
                    throw "App Token has expired"
                }

             }
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $AccessToken.access_token);
             $ClientResult = $HttpClient.GetAsync($RequestURL)
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {
                     Write-Output ("Error making REST Get " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
               return $JsonObject
             }
  
         }    
}

function Invoke-RestPOST
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [PSCustomObject]$AccessToken,
        [Parameter(Position=4, Mandatory=$true)] [PSCustomObject]$Content
    )  
 	Begin
		 {
             #Check for expired Token
             $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0,[System.DateTimeKind]::Utc);
             $expiry =  $minTime.AddSeconds($AccessToken.expires_on)
             if($expiry -le [DateTime]::Now.ToUniversalTime()){
                if([bool]($AccessToken.PSobject.Properties.name -match "refresh_token")){
                    write-host "Refresh Token"
                    $AccessToken = Refresh-AccessToken -MailboxName $MailboxName -RefreshToken $AccessToken.refresh_token               
                    Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
                }
                else{
                    throw "App Token has expired a new access token is required rerun get-apptoken"
                }
             }
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $AccessToken.access_token);
             $PostContent = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
             $ClientResult = $HttpClient.PostAsync([Uri]($RequestURL),$PostContent)
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {
                     Write-Output $ClientResult
                     Write-Output ("Error making REST POST " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
               return $JsonObject
             }
  
         }    
}

function Invoke-RestPatch
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [PSCustomObject]$AccessToken,
        [Parameter(Position=4, Mandatory=$true)] [PSCustomObject]$Content
    )  
 	Begin
		 {
             #Check for expired Token
             $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0,[System.DateTimeKind]::Utc);
             $expiry =  $minTime.AddSeconds($AccessToken.expires_on)
             if($expiry -le [DateTime]::Now.ToUniversalTime()){
                write-host "Refresh Token"
                $AccessToken = Refresh-AccessToken -MailboxName $MailboxName -RefreshToken $AccessToken.refresh_token               
                Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
             }
             $method =  New-Object System.Net.Http.HttpMethod("PATCH")
             $HttpRequestMessage =  New-Object System.Net.Http.HttpRequestMessage($method,[Uri]$RequestURL)
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $AccessToken.access_token);
             $HttpRequestMessage.Content = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
             $ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {
                     Write-Output $ClientResult
                     Write-Output ("Error making REST PATCH " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
               return $JsonObject
             }
  
         }    
}
function Invoke-RestDELETE
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [PSCustomObject]$AccessToken

    )  
 	Begin
		 {
             #Check for expired Token
             $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0,[System.DateTimeKind]::Utc);
             $expiry =  $minTime.AddSeconds($AccessToken.expires_on)
             if($expiry -le [DateTime]::Now.ToUniversalTime()){
                write-host "Refresh Token"
                $AccessToken = Refresh-AccessToken -MailboxName $MailboxName -RefreshToken $AccessToken.refresh_token               
                Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
             }
             $method =  New-Object System.Net.Http.HttpMethod("DELETE")
             $HttpRequestMessage =  New-Object System.Net.Http.HttpRequestMessage($method,[Uri]$RequestURL)
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $AccessToken.access_token);
             $ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {
                     Write-Output $ClientResult
                     Write-Output ("Error making REST DELETE " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               
               return $ClientResult.Result
             }
  
         }    
}
function Get-MailboxSettings{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-AutomaticRepliesSettings{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName  
                    
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings/AutomaticRepliesSetting"
       return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-MailboxTimeZone{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings/TimeZone"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-FolderFromPath{
	param (
			[Parameter(Position=0, Mandatory=$true)] [string]$FolderPath,
			[Parameter(Position=1, Mandatory=$true)] [string]$MailboxName,
            [Parameter(Position=2, Mandatory=$false)] [PSCustomObject]$AccessToken
		  )
	process{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders/msgfolderroot/childfolders?"
      #  $RootFolder = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		#Split the Search path into an array  
        $tfTargetFolder = $RootFolder
		$fldArray = $FolderPath.Split("\") 
		 #Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
	        #Perform search based on the displayname of each folder level
            $FolderName = $fldArray[$lint];
            $RequestURL = $RequestURL += "`$filter=DisplayName eq '$FolderName'" 
            $tfTargetFolder = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            if($tfTargetFolder.Value.Count -eq 1){
                $folderId = $tfTargetFolder.Value[0].Id.ToString()
                $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders('$folderId')/childfolders?"
            }
            else{
			    throw ("Folder Not found")
		    }
	    }  
		if($tfTargetFolder.Value.Count -gt 0){
            return ,$tfTargetFolder.Value[0]
		}
		else{
			throw ("Folder Not found")
		}
	}
}

function Get-Inbox{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders/Inbox"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-InboxItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders/Inbox/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,InferenceClassification`&`$Top=1000"
        do{
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                Write-Output $Message
            }           
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while(![String]::IsNullOrEmpty($RequestURL))       

    }
}

function Get-FocusedInboxItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders/Inbox/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,InferenceClassification`&`$Top=1000`&`$filter=InferenceClassification eq 'Focused'"
        do{
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                Write-Output $Message
            }           
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while(![String]::IsNullOrEmpty($RequestURL))       

    }
}

function Get-CalendarItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/events/?`$select=Start,End,Subject,Organizer`&`$Top=1000"
        do{
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                Write-Output $Message
            }           
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while(![String]::IsNullOrEmpty($RequestURL))       

    }
}

function Get-FolderItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [string]$FolderPath,
        [Parameter(Position=2, Mandatory=$false)] [PSCustomObject]$Folder
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        if($FolderPath -ne $null)
        {
            $Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName         
        }        
        if($Folder -ne $null)
        {
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  $Folder.'@odata.id' + "/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead`&`$Top=1000"
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Message in $JSONOutput.Value) {
                    Write-Output $Message
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
       } 
   

    }
}

function New-Folder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [string]$ParentFolderPath,
        [Parameter(Position=3, Mandatory=$true)] [string]$DisplayName

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $ParentFolder = Get-FolderFromPath -FolderPath $ParentFolderPath -AccessToken $AccessToken -MailboxName $MailboxName            
        if($ParentFolder  -ne $null)
        {
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  $ParentFolder.'@odata.id' + "/childfolders"
            $NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"}"
            write-host $NewFolderPost
            return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost

       } 
   

    }
}
function New-ContactFolder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=3, Mandatory=$true)] [string]$DisplayName

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/ContactFolders"
        $NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"}"
        write-host $NewFolderPost
        return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
   

    }
}
function New-CalendarFolder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=3, Mandatory=$true)] [string]$DisplayName

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/calendars"
        $NewFolderPost = "{`"Name`": `"" + $DisplayName + "`"}"
        write-host $NewFolderPost
        return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
   

    }
}

function Rename-Folder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [string]$FolderPath,
        [Parameter(Position=3, Mandatory=$true)] [string]$NewDisplayName

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName            
        if($Folder  -ne $null)
        {
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  $Folder.'@odata.id'
            $RenameFolderPost = "{`"DisplayName`": `"" + $NewDisplayName + "`"}"
            return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $RenameFolderPost

       } 
   

    }
}
function Update-FolderClass{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [string]$FolderPath,
        [Parameter(Position=3, Mandatory=$true)] [string]$FolderClass

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName            
        if($Folder  -ne $null)
        {
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  $Folder.'@odata.id'
            $UpdateFolderPost = "{`"SingleValueExtendedProperties`": [{`"PropertyId`":`"String 0x3613`",`"Value`":`"" + $FolderClass + "`"}]}"
            return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $UpdateFolderPost

       } 
   

    }
}
function Update-Folder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [string]$FolderPath,
        [Parameter(Position=3, Mandatory=$true)] [string]$FolderPost

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName            
        if($Folder  -ne $null)
        {
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  $Folder.'@odata.id'
            $FolderPostValue = $FolderPost
            return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $FolderPostValue

       } 
   

    }
}

function GetFolderRetentionTags(){
        #PR_POLICY_TAG 0x3019
    $PR_POLICY_TAG = Get-TaggedProperty -DataType "Binary" -Id "0x3019"  
    #PR_RETENTION_FLAGS 0x301D   
    $PR_RETENTION_FLAGS =  Get-TaggedProperty -DataType "Integer" -Id "0x301D" 
    #PR_RETENTION_PERIOD 0x301A
    $PR_RETENTION_PERIOD = Get-TaggedProperty -DataType "Integer" -Id "0x301A"    
}
function Set-FolderRetentionTag {
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [string]$FolderPath,
      	[Parameter(Position=3, Mandatory=$true)] [String]$PolicyTagValue,
		[Parameter(Position=4, Mandatory=$true)] [Int32]$RetentionFlagsValue,		
		[Parameter(Position=5, Mandatory=$true)] [Int32]$RetentionPeriodValue
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName            
        if($Folder  -ne $null)
        {    
 
            $retentionTagGUID = "{$($PolicyTagValue)}"
		    $policyTagGUID = new-Object Guid($retentionTagGUID) 
            $PolicyTagBase64 = [System.Convert]::ToBase64String($PolicyTagGUID.ToByteArray()) 
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  $Folder.'@odata.id'
            $FolderPostValue = "{`"SingleValueExtendedProperties`": [`r`n"
            $FolderPostValue += "`t{`"PropertyId`":`"Binary 0x3019`",`"Value`":`"" + $PolicyTagBase64 + "`"},`r`n"
            $FolderPostValue += "`t{`"PropertyId`":`"Integer 0x301D`",`"Value`":`"" + $RetentionFlagsValue + "`"},`r`n"
            $FolderPostValue += "`t{`"PropertyId`":`"Integer 0x301A`",`"Value`":`"" + $RetentionPeriodValue+ "`"}`r`n"
            $FolderPostValue += "]}"
            Write-Host $FolderPostValue
            return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $FolderPostValue
       } 
    }
    
}
function Delete-Folder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [string]$FolderPath
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName            
        if($Folder  -ne $null)
        {
            $confirmation = Read-Host "Are you Sure You Want To proceed with deleting Folder"
            if ($confirmation -eq 'y') {
                $HttpClient =  Get-HTTPClient($MailboxName)
                $RequestURL =  $Folder.'@odata.id'
                return Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            }
            else
            {
                Write-Host "skipped deletion"                
            }
       } 
   

    }
}

function Get-AllMailboxItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders/AllItems/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,ParentFolderId`&`$Top=1000"
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Message in $JSONOutput.Value) {
                    Write-Output $Message
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
   

    }
}

function GetExtendedPropList{
        param( 
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$PropertyList
    )
    Begin{
        $rtString = "";
        foreach($Prop in $PropertyList){
            if($Prop.PropertyType -eq "Tagged"){
                if($rtString -eq ""){
                    $rtString = "(PropertyId%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"                 
                }
                else{
                    $rtString += " or (PropertyId%20eq%20'" + $Prop.DataType + "%20" + $Prop.Id + "')"       
                }
            }
            else{
                if($Prop.Type -eq "String"){
                    if($rtString -eq ""){
                        $rtString = "(PropertyId%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20"+ $Prop.Id + "')"                 
                    }
                    else{
                        $rtString += " or (PropertyId%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Name%20"+ $Prop.Id + "')"       
                    }   
                }
                else{
                    if($rtString -eq ""){
                        $rtString = "(PropertyId%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20"+ $Prop.Id + "')"                 
                    }
                    else{
                        $rtString += " or (PropertyId%20eq%20'" + $Prop.DataType + "%20{" + $Prop.Guid + "}%20Id%20"+ $Prop.Id + "')"       
                    }                    
                }
            }
            
        }
        return $rtString
        
    }
}

function Get-TaggedProperty{
     param( 
        [Parameter(Position=0, Mandatory=$true)] [String]$DataType,
        [Parameter(Position=1, Mandatory=$true)] [String]$Id
    )
    Begin{
        $Property = "" | Select Id,DataType,PropertyType
        $Property.Id = $Id
        $Property.DataType = $DataType   
        $Property.PropertyType = "Tagged"
        return ,$Property
    }
}

function Get-NamedProperty{
     param( 
        [Parameter(Position=0, Mandatory=$true)] [String]$DataType,
        [Parameter(Position=1, Mandatory=$true)] [String]$Id,
        [Parameter(Position=1, Mandatory=$true)] [String]$Guid,
        [Parameter(Position=1, Mandatory=$true)] [String]$Type
    )
    Begin{
        $Property = "" | Select Id,DataType,PropertyType,Guid,$Type
        $Property.Id = $Id
        $Property.DataType = $DataType   
        $Property.PropertyType = "Named"
        if($Type = "String"){
            $Property.Type = "String"
        }
        else{
             $Property.Type = "Id"
        }
        return ,$Property
    }
}

function Get-FolderClass()
{
    $FolderClass = "" | Select Id,DataType,PropertyType
    $FolderClass.Id = "0x3613"
    $FolderClass.DataType = "String"  
    $FolderClass.PropertyType = "Tagged" 
    return ,$FolderClass
}
function Get-FolderPath()
{
    $FolderPath = "" | Select Id,DataType,PropertyType
    $FolderPath.Id = "0x66B5"
    $FolderPath.DataType = "String"
    $FolderPath.PropertyType = "Tagged"   
    return ,$FolderPath
}

function Get-AllMailFolders{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [PSCustomObject]$PropList
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders/msgfolderroot/childfolders/?`$Top=1000"
            if($PropList -ne $null){
               $Props = GetExtendedPropList -PropertyList $PropList
               $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
               Write-Host $RequestURL
            }
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Folder in $JSONOutput.Value) {
                    $Folder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ("\\" + $Folder.DisplayName)
                    Write-Output $Folder
                    if($Folder.ChildFolderCount -gt 0)
                    {
                        if($PropList -ne $null){
                            Get-AllChildFolders -Folder $Folder -AccessToken $AccessToken -PropList $PropList     
                        }
                        else{                            
                             Get-AllChildFolders -Folder $Folder -AccessToken $AccessToken     
                        }                                           
                    }
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
   

    }
}
function Get-AllChildFolders{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [PSCustomObject]$Folder,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [PSCustomObject]$PropList
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =   $Folder.'@odata.id' + "/childfolders/?`$Top=1000"
            if($PropList -ne $null){
               $Props = GetExtendedPropList -PropertyList $PropList
               $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($ChildFolder in $JSONOutput.Value) {
                    $ChildFolder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ($Folder.FolderPath + "\" + $ChildFolder.DisplayName)
                    Write-Output $ChildFolder
                    if($ChildFolder.ChildFolderCount -gt 0)
                    {
                        if($PropList -ne $null){
                            Get-AllChildFolders -Folder $ChildFolder -AccessToken $AccessToken -PropList $PropList     
                        }
                        else{                            
                            Get-AllChildFolders -Folder $ChildFolder -AccessToken $AccessToken     
                        }                
                    }
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
   

    }
}

function Get-AllCalendarFolders{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/Calendars/?`$Top=1000`&`$expand=SingleValueExtendedProperties(`$filter=PropertyId%20eq%20'String%200x66B5')"
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Folder in $JSONOutput.Value) {
                    Write-Output $Folder
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
   

    }
}
function Get-AllContactFolders{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/contactfolders/?`$Top=1000`&`$expand=SingleValueExtendedProperties(`$filter=PropertyId%20eq%20'String%200x66B5')"
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Folder in $JSONOutput.Value) {
                    Write-Output $Folder
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
   

    }
}
function Get-AllTaskfolders{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/taskfolders/?`$Top=1000"
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Folder in $JSONOutput.Value) {
                    Write-Output $Folder
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
   

    }
}



function Get-ArchiveFolder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings/ArchiveFolder"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        $folderId = $JsonObject.value.ToString()
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailFolders('$folderId')"
        return  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}


function Get-MailboxSettingsReport{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [psobject]$Mailboxes,
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Mandatory=$True)][Security.SecureString]$password   
    )
    Begin{
        $rptCollection = @()
        $AccessToken = Get-AppOnlyToken -CertFileName $CertFileName -password $password 
        $HttpClient =  Get-HTTPClient($Mailboxes[0])
        foreach ($MailboxName in $Mailboxes) {
            $rptObj = "" | Select MailboxName,Language,Locale,TimeZone,AutomaticReplyStatus
            $RequestURL =  "https://outlook.office.com/api/v2.0/Users('$MailboxName')/MailboxSettings"
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

function  Get-People {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/beta/me/people/?`$top=1000&`$Select=DisplayName"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        Write-Output $JsonObject 
    }
}


function  Get-MailboxUser {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/users/" + $MailboxName 
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        Write-Output $JsonObject 
    }
}

function  Get-CalendarGroups {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/users/" + $MailboxName + "/CalendarGroups"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        Write-Output $JsonObject 
    }
}

function  Enum-CalendarGroups {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [PSCustomObject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  "https://outlook.office.com/api/v2.0/users/" + $MailboxName + "/CalendarGroups"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        foreach($Group in $JsonObject.Value)
        {
            Write-Host ("GroupName : " + $Group.Name) 
            $GroupId = $Group.Id.ToString()           
            $RequestURL =  "https://outlook.office.com/api/v2.0/users/" + $MailboxName + "/CalendarGroups('$GroupId')/Calendars"
            $JsonObjectSub =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Calendar in $JsonObjectSub.Value) {
                Write-Host $Calendar.Name
            }
            $RequestURL =  "https://outlook.office.com/api/v2.0/users/" + $MailboxName + "/CalendarGroups('$GroupId')/MailFolders"
            $JsonObjectSub =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Calendar in $JsonObjectSub.Value) {
                Write-Host $Calendar.Name
            }
         
        }  
        
        
    }
}