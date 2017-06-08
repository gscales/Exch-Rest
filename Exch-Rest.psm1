##  A PowerShell module for the Office 365 and Exchange 2016 REST API
##  for documenation please refer to https://github.com/gscales/Exch-Rest
##
## The MIT License (MIT)
##
## Copyright (c) 2017 Glen Scales
##
## Permission is hereby granted, free of charge, to any person obtaining a copy
## of this software and associated documentation files (the "Software"), to deal
## in the Software without restriction, including without limitation the rights
## to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
## copies of the Software, and to permit persons to whom the Software is
## furnished to do so, subject to the following conditions:

## The above copyright notice and this permission notice shall be included in all
## copies or substantial portions of the Software.

## THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
## IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
## FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
## AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
## LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
## OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
## SOFTWARE.
function Get-AppSettings(){
        param( 
        
        )  
 	Begin
		 {
            $configObj = "" |  select ResourceURL,ClientId,redirectUrl,ClientSecret,x5t,TenantId,ValidateForMinutes
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

function Invoke-DecodeToken { 
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
            $x5t = [System.Convert]::ToBase64String($cert.GetCertHash())
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
        [Parameter(Position=3, Mandatory=$false)] [string]$ClientSecret,
        [Parameter(Position=4, Mandatory=$false)] [string]$ResourceURL
    )  
 	Begin
		 {
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient($MailboxName)
            $AppSetting = Get-AppSettings 
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
            if([String]::IsNullOrEmpty($ResourceURL)){
                $ResourceURL = $AppSetting.ResourceURL
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
            if([bool]($JsonObject.PSobject.Properties.name -match "refresh_token")){
                $JsonObject.refresh_token =  $JsonObject.refresh_token | ConvertTo-SecureString -AsPlainText -Force
            }
            if([bool]($JsonObject.PSobject.Properties.name -match "access_token")){
                $JsonObject.access_token =  $JsonObject.access_token | ConvertTo-SecureString -AsPlainText -Force
            }
            Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
            Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
            return $JsonObject
         }
}

function Get-AccessTokenUserAndPass{ 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [string]$ClientId,
        [Parameter(Position=2, Mandatory=$false)] [string]$ResourceURL,
        [Parameter(Position=3, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
		 {
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient($MailboxName)
            $AppSetting = Get-AppSettings 
           if($ClientId -eq $null){
                 $ClientId = $AppSetting.ClientId
            }      

            if([String]::IsNullOrEmpty($ResourceURL)){
                $ResourceURL = $AppSetting.ResourceURL
            }  
            $UserName = $Credentials.UserName.ToString()
            $password = $Credentials.GetNetworkCredential().password.ToString()    
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&grant_type=password&username=$username&password=$password"
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/common/oauth2/token"),$content)
            $JsonObject = ConvertFrom-Json -InputObject  $ClientReesult.Result.Content.ReadAsStringAsync().Result
            Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
            if([bool]($JsonObject.PSobject.Properties.name -match "refresh_token")){
                $JsonObject.refresh_token =  $JsonObject.refresh_token | ConvertTo-SecureString -AsPlainText -Force
            }
            if([bool]($JsonObject.PSobject.Properties.name -match "access_token")){
                $JsonObject.access_token =  $JsonObject.access_token | ConvertTo-SecureString -AsPlainText -Force
            }
            #Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
            return $JsonObject
         }
}



function Get-AppOnlyToken{ 
    param( 
       
        [Parameter(Position=1, Mandatory=$true)] [string]$CertFileName,
        [Parameter(Position=2, Mandatory=$false)] [string]$TenantId,
        [Parameter(Position=3, Mandatory=$false)] [string]$ClientId,
        [Parameter(Position=4, Mandatory=$false)] [string]$redirectUrl,     
        [Parameter(Position=6, Mandatory=$false)] [Int32]$ValidateForMinutes,
        [Parameter(Position=7, Mandatory=$false)] [string]$ResourceURL,
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
            if($redirectUrl -eq $null){
                $redirectUrl = $AppSetting.redirectUrl
            }
            if($ValidateForMinutes -eq 0){
                $ValidateForMinutes = $AppSetting.ValidateForMinutes               
            }
            if([String]::IsNullOrEmpty($ResourceURL)){
                $ResourceURL = $AppSetting.ResourceURL
            }        
            $JWTToken = New-JWTToken -CertFileName $CertFileName -password $password -TenantId $TenantId -ClientId $ClientId  -ValidateForMinutes $ValidateForMinutes
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient(" ")          
            $AuthorizationPostRequest = "resource=https%3A%2F%2F$ResourceURL&client_id=$ClientId&client_assertion_type=urn%3Aietf%3Aparams%3Aoauth%3Aclient-assertion-type%3Ajwt-bearer&client_assertion=$JWTToken&grant_type=client_credentials&redirect_uri=$redirectUrl"
            $content = New-Object System.Net.Http.StringContent($AuthorizationPostRequest, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded")
            $ClientReesult = $HttpClient.PostAsync([Uri]("https://login.windows.net/" + $TenantId + "/oauth2/token"),$content)
            $JsonObject = ConvertFrom-Json -InputObject  $ClientReesult.Result.Content.ReadAsStringAsync().Result
            if([bool]($JsonObject.PSobject.Properties.name -match "refresh_token")){
                $JsonObject.refresh_token =  $JsonObject.refresh_token | ConvertTo-SecureString -AsPlainText -Force
            }
            if([bool]($JsonObject.PSobject.Properties.name -match "access_token")){
                $JsonObject.access_token =  $JsonObject.access_token | ConvertTo-SecureString -AsPlainText -Force
            }
            Add-Member -InputObject $JsonObject -NotePropertyName tenantid -NotePropertyValue $TenantId
            Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $ClientId
            Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
            return $JsonObject
         }
}



function Invoke-RefreshAccessToken{ 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [psobject]$AccessToken
    )  
 	Begin
		 {
            Add-Type -AssemblyName System.Web
            $HttpClient =  Get-HTTPClient($MailboxName)
            $ClientId = $AccessToken.clientid            
           # $redirectUrl = [System.Web.HttpUtility]::UrlEncode($AccessToken.redirectUrl)
            $redirectUrl = $AccessToken.redirectUrl
            $RefreshToken = (Get-TokenFromSecureString -SecureToken $AccessToken.refresh_token)
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
               Add-Member -InputObject $JsonObject -NotePropertyName clientid -NotePropertyValue $AccessToken.clientid
               Add-Member -InputObject $JsonObject -NotePropertyName redirectUrl -NotePropertyValue $AccessToken.redirectUrl
               return $JsonObject
             }

         }
}


function Get-TokenFromSecureString{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [System.Security.SecureString]$SecureToken
    )
    begin{
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
        $Token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        return,$Token
    }
}
function Invoke-RestGet
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [psobject]$AccessToken,
        [Parameter(Position=4, Mandatory=$false)] [switch]$NoJSON
    )  
 	Begin
		 {
             #Check for expired Token
             $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0,[System.DateTimeKind]::Utc);
             $expiry =  $minTime.AddSeconds($AccessToken.expires_on)
             if($expiry -le [DateTime]::Now.ToUniversalTime()){
                if([bool]($AccessToken.PSobject.Properties.name -match "refresh_token")){
                    write-host "Refresh Token"
                    $AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken          
                    Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
                }
                else{
                    throw "App Token has expired"
                }

             }
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
             $HttpClient.DefaultRequestHeaders.Add("Prefer", ("outlook.timezone=`"" + [TimeZoneInfo]::Local.Id + "`"")) 
             $ClientResult = $HttpClient.GetAsync($RequestURL)
             if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK){
                 if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created){
                     write-Host ($ClientResult.Result)
                 }
             }             
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
               if($NoJSON){
                    return  $ClientResult.Result.Content  
               }
               else{
                   $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
                   if([String]::IsNullOrEmpty($ClientResult)){
                        write-host "No Value returned"
                   }
                   else{
                       return $JsonObject
                   }

               }  

             }
  
         }    
}

function Invoke-RestPOST
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [psobject]$AccessToken,
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
                    $AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken                
                    Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
                }
                else{
                    throw "App Token has expired a new access token is required rerun get-apptoken"
                }
             }
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
             $PostContent = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
             $HttpClient.DefaultRequestHeaders.Add("Prefer", ("outlook.timezone=`"" + [TimeZoneInfo]::Local.Id + "`"")) 
             $ClientResult = $HttpClient.PostAsync([Uri]($RequestURL),$PostContent)
             if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK){
                 if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created){
                     write-Host ($ClientResult.Result)
                 }
             }   
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
               if([String]::IsNullOrEmpty($JsonObject)){
                   Write-Output $ClientResult.Result
               }
               else{
                   return $JsonObject
               }
               
             }
  
         }    
}

function Invoke-RestPatch
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [psobject]$AccessToken,
        [Parameter(Position=4, Mandatory=$true)] [PSCustomObject]$Content
    )  
 	Begin
		 {
             #Check for expired Token
             $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0,[System.DateTimeKind]::Utc);
             $expiry =  $minTime.AddSeconds($AccessToken.expires_on)
             if($expiry -le [DateTime]::Now.ToUniversalTime()){
                write-host "Refresh Token"
                $AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken            
                Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
             }
             $method =  New-Object System.Net.Http.HttpMethod("PATCH")
             $HttpRequestMessage =  New-Object System.Net.Http.HttpRequestMessage($method,[Uri]$RequestURL)
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
             $HttpRequestMessage.Content = New-Object System.Net.Http.StringContent($Content, [System.Text.Encoding]::UTF8, "application/json")
             $ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
             if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK){
                 if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::Created){
                     write-Host ($ClientResult.Result)
                 }
             }   
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {
                     Write-Output ("Error making REST Patch " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
               if([String]::IsNullOrEmpty($JsonObject)){
                   Write-Output $ClientResult.Result
               }
               else{
                   return $JsonObject
               }
               
             }
  
         }    
}
function Invoke-RestDELETE
{
        param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$RequestURL,
        [Parameter(Position=1, Mandatory=$true)] [String]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)] [System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Position=3, Mandatory=$true)] [psobject]$AccessToken

    )  
 	Begin
		 {
             #Check for expired Token
             $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0,[System.DateTimeKind]::Utc);
             $expiry =  $minTime.AddSeconds($AccessToken.expires_on)
             if($expiry -le [DateTime]::Now.ToUniversalTime()){
                write-host "Refresh Token"
                $AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken          
                Set-Variable -Name "AccessToken" -Value $AccessToken -Scope Script -Visibility Public
             }
             $method =  New-Object System.Net.Http.HttpMethod("DELETE")
             $HttpRequestMessage =  New-Object System.Net.Http.HttpRequestMessage($method,[Uri]$RequestURL)
             $HttpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", (Get-TokenFromSecureString -SecureToken $AccessToken.access_token));
             $ClientResult = $HttpClient.SendAsync($HttpRequestMessage)
             if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::OK){
                 if($ClientResult.Result.StatusCode -ne [System.Net.HttpStatusCode]::NoContent){
                     write-Host ($ClientResult.Result)
                 }
             }   
             if (!$ClientResult.Result.IsSuccessStatusCode)
             {
                     Write-Output ("Error making REST Delete " + $ClientResult.Result.StatusCode + " : " + $ClientResult.Result.ReasonPhrase)
                     Write-Output $ClientResult.Result
                     if($ClientResult.Content -ne $null){
                         Write-Output ($ClientResult.Content.ReadAsStringAsync().Result);   
                     }                     
             }
            else
             {
               $JsonObject = ConvertFrom-Json -InputObject  $ClientResult.Result.Content.ReadAsStringAsync().Result
               if([String]::IsNullOrEmpty($JsonObject)){
                   Write-Output $ClientResult.Result
               }
               else{
                   return $JsonObject
               }
               
             }
  
         }    
}
function Get-MailboxSettings{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('$MailboxName')/MailboxSettings"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-AutomaticRepliesSettings{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName  
                    
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings/AutomaticRepliesSetting"
       return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-MailboxTimeZone{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('$MailboxName')/MailboxSettings/TimeZone"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-FolderFromPath{
	param (
			[Parameter(Position=0, Mandatory=$true)] [string]$FolderPath,
			[Parameter(Position=1, Mandatory=$true)] [string]$MailboxName,
            [Parameter(Position=2, Mandatory=$false)] [psobject]$AccessToken
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
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('$MailboxName')/MailFolders/msgfolderroot/childfolders?"
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
                $RequestURL =  $EndPoint + "('$MailboxName')/MailFolders('$folderId')/childfolders?"
            }
            else{
			    throw ("Folder Not found")
		    }
	    }  
		if($tfTargetFolder.Value.Count -gt 0){
            $folderId = $tfTargetFolder.Value[0].Id.ToString()
            Add-Member -InputObject $tfTargetFolder.Value[0] -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =   $EndPoint + "('$MailboxName')/MailFolders/Inbox"
        return Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
    }
}

function Get-InboxItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders/Inbox/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,InferenceClassification`&`$Top=1000"
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('$MailboxName')/MailFolders/Inbox/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,InferenceClassification`&`$Top=1000`&`$filter=InferenceClassification eq 'Focused'"
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }   
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('$MailboxName')/events/?`$select=Start,End,Subject,Organizer`&`$Top=1000"
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [string]$FolderPath,
        [Parameter(Position=3, Mandatory=$false)] [PSCustomObject]$Folder,
        [Parameter(Position=4, Mandatory=$false)] [switch]$ReturnSize,
        [Parameter(Position=5, Mandatory=$false)] [string]$SelectProperties,
        [Parameter(Position=6, Mandatory=$false)] [string]$Filter
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
        if(![String]::IsNullorEmpty($Filter)){
            $Filter = "`&`$filter=" + $Filter
        }        
        if([String]::IsNullorEmpty($SelectProperties)){
            $SelectProperties = "`$select=ReceivedDateTime,Sender,Subject,IsRead"
        }
        else{
            $SelectProperties = "`$select=" + $SelectProperties
        }
        if($Folder -ne $null)
        {
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  $Folder.FolderRestURI + "/messages/?" +  $SelectProperties + "`&`$Top=1000" + $Filter
            write-host $RequestURL
            if($ReturnSize.IsPresent){
                $RequestURL =  $Folder.FolderRestURI + "/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead`&`$Top=1000`&`$expand=SingleValueExtendedProperties(`$filter=PropertyId%20eq%20'Integer%200x0E08')" + $Filter
            }
           
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Message in $JSONOutput.Value) {
                    Add-Member -InputObject $Message -NotePropertyName ItemRESTURI -NotePropertyValue ($Folder.FolderRestURI + "/messages('" + $Message.Id + "')")
                    Write-Output $Message
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL))     
       } 
   

    }
}

function Move-Message{
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$ItemURI,
        [Parameter(Position=2, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=3, Mandatory=$false)] [string]$TargetFolderPath
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }    
        if($TargetFolderPath -ne $null)
        {
            $Folder = Get-FolderFromPath -FolderPath $TargetFolderPath -AccessToken $AccessToken -MailboxName $MailboxName         
        }
        if($Folder -ne $null){
            $HttpClient =  Get-HTTPClient($MailboxName)
            $RequestURL =  $ItemURI + "/move"
            $MoveItemPost = "{`"DestinationId`": `"" + $Folder.Id + "`"}"
            write-host $MoveItemPost
            return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $MoveItemPost
        } 
    }
}
function Update-Message{
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$ItemURI,
        [Parameter(Position=2, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=3, Mandatory=$false)] [String]$Subject,
        [Parameter(Position=4, Mandatory=$false)] [String]$Body,
        [Parameter(Position=5, Mandatory=$false)] [psobject]$Attachments,
        [Parameter(Position=6, Mandatory=$false)] [psobject]$ToRecipients,
        [Parameter(Position=7, Mandatory=$false)] [psobject]$StandardPropList,
        [Parameter(Position=8, Mandatory=$false)] [psobject]$ExPropList

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }    
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  $ItemURI
        $UpdateItemPatch =   Get-MessageJSONFormat -Subject $Subject -Body $Body -Attachments $Attachments -ExPropList $ExPropList -StandardPropList $StandardPropList  
        Write-host $UpdateItemPatch
        return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $UpdateItemPatch
    }
}


function Get-Attachments{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$ItemURI,
        [Parameter(Position=2, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=3, Mandatory=$false)] [switch]$MetaData,
        [Parameter(Position=4, Mandatory=$false)] [string]$SelectProperties
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }    
        if([String]::IsNullorEmpty($SelectProperties)){
            $SelectProperties = "`$select=Name,ContentType,Size,isInline,ContentType"
        }
        else{
            $SelectProperties = "`$select=" + $SelectProperties
        }
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  $ItemURI + "/Attachments?" +  $SelectProperties   
        do{
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                  Add-Member -InputObject $Message -NotePropertyName AttachmentRESTURI -NotePropertyValue ($ItemURI + "/Attachments('" + $Message.Id + "')")
                  Write-Output $Message
             }           
             $RequestURL = $JSONOutput.'@odata.nextLink'
         }while(![String]::IsNullOrEmpty($RequestURL))     
   

    }
}

function Invoke-DownloadAttachment{
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$AttachmentURI,
        [Parameter(Position=2, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        } 
        $HttpClient =  Get-HTTPClient($MailboxName)
        $AttachmentURI = $AttachmentURI + "?`$expand"
        $AttachmentObj = Invoke-RestGet -RequestURL $AttachmentURI -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName   
        return $AttachmentObj
    }
} 



function New-Folder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
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
            $RequestURL =  $ParentFolder.FolderRestURI + "/childfolders"
            $NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"}"
            write-host $NewFolderPost
            return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost

       } 
   

    }
}
function New-ContactFolder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=3, Mandatory=$true)] [string]$DisplayName

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('$MailboxName')/ContactFolders"
        $NewFolderPost = "{`"DisplayName`": `"" + $DisplayName + "`"}"
        write-host $NewFolderPost
        return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
   

    }
}
function New-CalendarFolder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=3, Mandatory=$true)] [string]$DisplayName

    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('$MailboxName')/calendars"
        $NewFolderPost = "{`"Name`": `"" + $DisplayName + "`"}"
        write-host $NewFolderPost
        return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewFolderPost
   

    }
}

function Rename-Folder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
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
            $RequestURL =  $Folder.FolderRestURI
            $RenameFolderPost = "{`"DisplayName`": `"" + $NewDisplayName + "`"}"
            return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $RenameFolderPost

       } 
   

    }
}
function Update-FolderClass{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
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
            $RequestURL =  $Folder.FolderRestURI
            $UpdateFolderPost = "{`"SingleValueExtendedProperties`": [{`"PropertyId`":`"String 0x3613`",`"Value`":`"" + $FolderClass + "`"}]}"
            return Invoke-RestPatch -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $UpdateFolderPost

       } 
   

    }
}
function Update-Folder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
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
            $RequestURL =  $Folder.FolderRestURI
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
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
            $RequestURL =  $Folder.FolderRestURI
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

function Invoke-DeleteItem{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$true)] [string]$ItemURI
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        $confirmation = Read-Host "Are you Sure You Want To proceed with deleting the Item"
        if ($confirmation -eq 'y') {
             $HttpClient =  Get-HTTPClient($MailboxName)
             $RequestURL =  $ItemURI
             return Invoke-RestDELETE -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        }
        else
        {
             Write-Host "skipped deletion"                
        }
   

    }
}

function Invoke-DeleteFolder{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
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
                $RequestURL =  $Folder.FolderRestURI
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =  $EndPoint + "('$MailboxName')/MailFolders/AllItems/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead,ParentFolderId`&`$Top=1000"
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
        [Parameter(Position=1, Mandatory=$true)] [String]$Id,
        [Parameter(Position=2, Mandatory=$false)] [String]$Value
    )
    Begin{
        $Property = "" | Select Id,DataType,PropertyType,Value
        $Property.Id = $Id
        $Property.DataType = $DataType   
        $Property.PropertyType = "Tagged"
        if(![String]::IsNullOrEmpty($Value)){
             $Property.Value = $Value
        }       
        return ,$Property
    }
}

function Get-NamedProperty{
     param( 
        [Parameter(Position=0, Mandatory=$true)] [String]$DataType,
        [Parameter(Position=1, Mandatory=$true)] [String]$Id,
        [Parameter(Position=1, Mandatory=$true)] [String]$Guid,
        [Parameter(Position=1, Mandatory=$true)] [String]$Type,
        [Parameter(Position=2, Mandatory=$false)] [String]$Value
    )
    Begin{
        $Property = "" | Select Id,DataType,PropertyType,Type,Guid,Value
        $Property.Id = $Id
        $Property.DataType = $DataType   
        $Property.PropertyType = "Named"
        $Property.Guid = $Guid
        if($Type = "String"){
            $Property.Type = "String"
        }
        else{
             $Property.Type = "Id"
        }
         if(![String]::IsNullOrEmpty($Value)){
             $Property.Value = $Value
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [PSCustomObject]$PropList
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =  $EndPoint + "('$MailboxName')/MailFolders/msgfolderroot/childfolders/?`$Top=1000"
            if($PropList -ne $null){
               $Props = GetExtendedPropList -PropertyList $PropList
               $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
               Write-Host $RequestURL
            }
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Folder in $JSONOutput.Value) {
                    $Folder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ("\\" + $Folder.DisplayName)
                    $folderId = $Folder.Id.ToString()
                    Add-Member -InputObject $Folder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [PSCustomObject]$PropList
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =   $Folder.FolderRestURI + "/childfolders/?`$Top=1000"
            if($PropList -ne $null){
               $Props = GetExtendedPropList -PropertyList $PropList
               $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($ChildFolder in $JSONOutput.Value) {
                    $ChildFolder | Add-Member -NotePropertyName FolderPath -NotePropertyValue ($Folder.FolderPath + "\" + $ChildFolder.DisplayName)
                    $folderId = $ChildFolder.Id.ToString()
                    Add-Member -InputObject $ChildFolder -NotePropertyName FolderRestURI -NotePropertyValue ($EndPoint + "('$MailboxName')/MailFolders('$folderId')")
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [switch]$FolderClass
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            if($FolderClass.IsPresent){
                $RequestURL =  $EndPoint + "('$MailboxName')/Calendars/?`$Top=1000`&`$expand=SingleValueExtendedProperties(`$filter=PropertyId%20eq%20'String%200x66B5')"
                if($AccessToken.resource -eq "https://graph.microsoft.com"){
                    $RequestURL =  $EndPoint + "('$MailboxName')/Calendars/?`$Top=1000`&`$expand=SingleValueExtendedProperties(`$filter=Id%20eq%20'String%200x66B5')"
                }
                
            }
            else{
                 $RequestURL =  $EndPoint + "('$MailboxName')/Calendars/?`$Top=1000"
            }
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =  $EndPoint + "('$MailboxName')/contactfolders/?`$Top=1000"
            Write-Host  $RequestURL
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
            $HttpClient =  Get-HTTPClient($MailboxName)
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =  $EndPoint + "('$MailboxName')/taskfolders/?`$Top=1000"
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
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "('$MailboxName')/MailboxSettings/ArchiveFolder"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        $folderId = $JsonObject.value.ToString()
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL = $EndPoint + "('$MailboxName')/MailFolders('$folderId')"
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
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =  $EndPoint + "('$MailboxName')/MailboxSettings"
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

function Get-EndPoint{
        param(
        [Parameter(Position=0, Mandatory=$true)] [psObject]$AccessToken,
        [Parameter(Position=1, Mandatory=$true)] [psObject]$Segment
    )
    Begin{
        $EndPoint = "https://outlook.office.com/api/v2.0"
        switch($AccessToken.resource){
            "https://outlook.office.com" {  $EndPoint = "https://outlook.office.com/api/v2.0/" + $Segment }     
            "https://graph.microsoft.com" {  $EndPoint = "https://graph.microsoft.com/v1.0/" + $Segment  }     
        }
        return , $EndPoint
        
    }
}

function  Get-People {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken   
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

function  Get-UserPhotoMetaData {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "/" + $MailboxName + "/photo"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        Write-Output $JsonObject 
    }
}

function  Get-UserPhoto {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken
 
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "/" + $MailboxName + "/photo/`$value"
        $Result =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -NoJSON
        Write-Output $Result.ReadAsByteArrayAsync().Result  
    }
}
function  Get-MailboxUser {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "/" + $MailboxName 
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        Write-Output $JsonObject 
    }
}

function  Get-CalendarGroups {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "/" + $MailboxName + "/CalendarGroups"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        Write-Output $JsonObject 
    }
}

function  Invoke-EnumCalendarGroups {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "/" + $MailboxName + "/CalendarGroups"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        foreach($Group in $JsonObject.Value)
        {
            Write-Host ("GroupName : " + $Group.Name) 
            $GroupId = $Group.Id.ToString()           
            $RequestURL =  $EndPoint + "/" + $MailboxName + "/CalendarGroups('$GroupId')/Calendars"
            $JsonObjectSub =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Calendar in $JsonObjectSub.Value) {
                Write-Host $Calendar.Name
            }
            $RequestURL =  $EndPoint + "/" + $MailboxName + "/CalendarGroups('$GroupId')/MailFolders"
            $JsonObjectSub =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Calendar in $JsonObjectSub.Value) {
                Write-Host $Calendar.Name
            }
         
        }  
        
        
    }
}

function Get-ObjectProp{
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$Name,  
        [Parameter(Position=1, Mandatory=$false)] [psObject]$PropList,
        [Parameter(Position=2, Mandatory=$false)] [switch]$Array              
    )
    Begin{
        $ObjectProp = "" | Select PropertyName,PropertyList,PropertyType,isArray
        $ObjectProp.PropertyType = "Object"
        $ObjectProp.isArray = $false
        if($Array.IsPresent){  $ObjectProp.isArray = $true }
        $ObjectProp.PropertyName = $Name
        if($PropList -eq $null){
            $ObjectProp.PropertyList = @()
        }
        else{
            $ObjectProp.PropertyList = $PropList
        }
        return ,$ObjectProp
        
    }
}

function Get-ObjectCollectionProp{
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$Name,  
        [Parameter(Position=1, Mandatory=$false)] [psObject]$PropList,
        [Parameter(Position=2, Mandatory=$false)] [switch]$Array              
    )
    Begin{
        $CollectionProp = "" | Select PropertyName,PropertyList,PropertyType,isArray
        $CollectionProp.PropertyType = "ObjectCollection"
        $CollectionProp.isArray = $false
        if($Array.IsPresent){  $CollectionProp.isArray = $true }
        $CollectionProp.PropertyName = $Name
        if($PropList -eq $null){
            $CollectionProp.PropertyList = @()
        }
        else{
            $CollectionProp.PropertyList = $PropList
        }
        return ,$CollectionProp
        
    }
}

function Get-ItemProp{
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$Name,  
        [Parameter(Position=1, Mandatory=$true)] [string]$Value,  
        [Parameter(Position=2, Mandatory=$false)] [switch]$NoQuotes          
    )
    Begin{
        $ItemProp = "" | Select Name,Value,PropertyType,QuoteValue
        $ItemProp.PropertyType = "Single"
        $ItemProp.Name = $Name
        $ItemProp.Value = $Value  
        if($NoQuotes.IsPresent){
            $ItemProp.QuoteValue = $false  
        }
        else{
            $ItemProp.QuoteValue = $true  
        }
        return ,$ItemProp
        
    }
}

function List-Groups {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken   
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-HTTPClient($MailboxName)
        $RequestURL =  Get-EndPoint -AccessToken $AccessToken -Segment "/groups"
        $JsonObject =  Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        return $JsonObject       
        
    }
}

function Get-MailAppProps(){
        #Holder for Mail Apps
        $cepPropdef = Get-NamedProperty -DataType "String" -Guid "00020329-0000-0000-C000-000000000046" -Id "cecp-propertyNames" -Value ($Guid + ";") -Type "String"
        $cepPropValue = Get-NamedProperty -DataType "String" -Guid "00020329-0000-0000-C000-000000000046" -Id "cecp-" + $Guid -Value $value -Type "String"
}

function New-EmailAddress {
    param(
        [Parameter(Position=0, Mandatory=$false)] [string]$Name,
        [Parameter(Position=1, Mandatory=$true)] [string]$Address
    )
    Begin{
        $EmailAddress = "" | Select Name,Address
        if([String]::IsNullOrEmpty($Name)){
            $EmailAddress.Name = $Address
        }
        else{
            $EmailAddress.Name = $Name
        }        
        $EmailAddress.Address = $Address
        return, $EmailAddress
    }
}

#region Sending_Email
function  New-SentEmailMessage {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [string]$FolderPath,
        [Parameter(Position=3, Mandatory=$false)] [PSCustomObject]$Folder,
        [Parameter(Position=4, Mandatory=$true)] [String]$Subject,
        [Parameter(Position=5, Mandatory=$false)] [String]$Body,
        [Parameter(Position=7, Mandatory=$true)] [psobject]$SenderEmailAddress,
        [Parameter(Position=8, Mandatory=$false)] [psobject]$Attachments,
        [Parameter(Position=9, Mandatory=$false)] [psobject]$ToRecipients,
        [Parameter(Position=10, Mandatory=$true)] [DateTime]$SentDate,
        [Parameter(Position=11, Mandatory=$false)] [psobject]$ExPropList,
        [Parameter(Position=12, Mandatory=$false)] [string]$ItemClass
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        } 
        $SentFlag = Get-TaggedProperty -DataType "Integer" -Id "0x0E07"  -Value "1"
        $SentTime = Get-TaggedProperty -DataType "SystemTime" -Id "0x0039"  -Value $SentDate.ToString("yyyy-MM-ddTHH:mm:ss.ffffzzz")
        $RcvdTime = Get-TaggedProperty -DataType "SystemTime" -Id "0x0E06"  -Value $SentDate.ToString("yyyy-MM-ddTHH:mm:ss.ffffzzz")
        if($ExPropList -eq $null){
            $ExPropList = @()
        }
        if(![String]::IsNullOrEmpty($ItemClass)){
            $ItemClassProp = Get-TaggedProperty -DataType "String" -Id "0x001A"  -Value $ItemClass
            $ExPropList += $ItemClassProp
        }        
        $ExPropList += $SentFlag
        $ExPropList += $SentTime
        $ExPropList += $RcvdTime
        $NewMessage = Get-MessageJSONFormat -Subject $Subject -Body $Body -SenderEmailAddress $SenderEmailAddress -Attachments $Attachments -ToRecipients $ToRecipients -SentDate $SentDate -ExPropList $ExPropList
        if($FolderPath -ne $null)
        {
            $Folder = Get-FolderFromPath -FolderPath $FolderPath -AccessToken $AccessToken -MailboxName $MailboxName    

        }        
        if($Folder -ne $null)
        {
            $RequestURL =  $Folder.FolderRestURI + "/messages"
            $HttpClient =  Get-HTTPClient($MailboxName)
            Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage
        }

    }
}

function CreateFlatList{
    param(
        [Parameter(Position=0, Mandatory=$true)] [psobject]$EmailAddress
    )
    Begin{
         
         $FlatListEntry = new-object System.IO.MemoryStream
         $EntryOneOffid = "00000000812B1FA4BEA310199D6E00DD010F540200000190" + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(($EmailAddress.Name + "`0"))).Replace("-","") + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(("SMTP" + "`0"))).Replace("-","") + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(($EmailAddress.Address + "`0"))).Replace("-","")
         $FlatListEntryBytes = HexStringToByteArray($EntryOneOffid)
         $FlatListEntry.Write([BitConverter]::GetBytes($FlatListEntryBytes.Length), 0, 4);
         $FlatListEntry.Write($FlatListEntryBytes, 0, $FlatListEntryBytes.Length);
         $InnerLength += $FlatListEntryBytes.Length
         $Modulsval = $FlatListEntryBytes.Length % 4;
         $PadingValue = 0;
         if ($Modulsval -ne 0)
         {
              $PadingValue = 4 - $Modulsval;
              for ($AddPading = 0; $AddPading -lt $PadingValue; $AddPading++)
              {
                   [Byte]$NullValue = 00;
                   $FlatlistStream.Write($NullValue, 0, 1);
              }
         }         
         $FlatListEntry.Position = 0
         $FlatListEntryBytes = $FlatListEntry.ToArray()
         $FlatList = new-object System.IO.MemoryStream
         $FlatList.Write([BitConverter]::GetBytes(1), 0, 4);  
         $FlatList.Write([BitConverter]::GetBytes($FlatListEntryBytes.Length), 0, 4); 
         $FlatList.Write($FlatListEntryBytes, 0,  $FlatListEntryBytes.Length);
         $FlatList.Position = 0
         return ,$FlatList.ToArray()   
     }
}


function Send-MessageREST{
        param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [string]$FolderPath,
        [Parameter(Position=3, Mandatory=$false)] [PSCustomObject]$Folder,
        [Parameter(Position=4, Mandatory=$true)] [String]$Subject,
        [Parameter(Position=5, Mandatory=$false)] [String]$Body,
        [Parameter(Position=7, Mandatory=$false)] [psobject]$SenderEmailAddress,
        [Parameter(Position=8, Mandatory=$false)] [psobject]$Attachments,
        [Parameter(Position=9, Mandatory=$false)] [psobject]$ToRecipients,
        [Parameter(Position=10, Mandatory=$false)] [psobject]$CCRecipients,
        [Parameter(Position=11, Mandatory=$false)] [psobject]$BCCRecipients,
        [Parameter(Position=12, Mandatory=$false)] [psobject]$ExPropList,
        [Parameter(Position=13, Mandatory=$false)] [psobject]$StandardPropList,
        [Parameter(Position=14, Mandatory=$false)] [string]$ItemClass,
        [Parameter(Position=15, Mandatory=$false)] [switch]$SaveToSentItems,
        [Parameter(Position=16, Mandatory=$false)] [switch]$ShowRequest,
        [Parameter(Position=17, Mandatory=$false)] [switch]$RequestReadRecipient,
        [Parameter(Position=18, Mandatory=$false)] [switch]$RequestDeliveryRecipient,
        [Parameter(Position=19, Mandatory=$false)] [psobject]$ReplyTo
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        } 
        if(![String]::IsNullOrEmpty($ItemClass)){
            $ItemClassProp = Get-TaggedProperty -DataType "String" -Id "0x001A"  -Value $ItemClass
            if($ExPropList -eq $null){
               $ExPropList = @()
            }
            $ExPropList += $ItemClassProp
        }     
        $SaveToSentFolder = "false"
        if($SaveToSentItems.IsPresent){
            $SaveToSentFolder = "true"
        }
        $NewMessage = Get-MessageJSONFormat -Subject $Subject -Body $Body -SenderEmailAddress $SenderEmailAddress -Attachments $Attachments -ToRecipients $ToRecipients -SentDate $SentDate -ExPropList $ExPropList -CcRecipients $CCRecipients -bccRecipients $BCCRecipients -StandardPropList  $StandardPropList -SaveToSentItems $SaveToSentFolder -SendMail -ReplyTo $ReplyTo -RequestReadRecipient $RequestReadRecipient.IsPresent -RequestDeliveryRecipient $RequestDeliveryRecipient.IsPresent
        if($ShowRequest.IsPresent){
            write-host $NewMessage
        }       
        $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "/" + $MailboxName + "/sendmail"
        $HttpClient =  Get-HTTPClient($MailboxName)
        return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage

    }
}

function Get-MessageJSONFormat {
    param(
        [Parameter(Position=1, Mandatory=$false)] [String]$Subject,
        [Parameter(Position=2, Mandatory=$false)] [String]$Body,
        [Parameter(Position=3, Mandatory=$false)] [psobject]$SenderEmailAddress,
        [Parameter(Position=5, Mandatory=$false)] [psobject]$Attachments,
        [Parameter(Position=6, Mandatory=$false)] [psobject]$ToRecipients,
        [Parameter(Position=7, Mandatory=$false)] [psobject]$CcRecipients,
        [Parameter(Position=7, Mandatory=$false)] [psobject]$bccRecipients,
        [Parameter(Position=8, Mandatory=$false)] [psobject]$SentDate,
        [Parameter(Position=9, Mandatory=$false)] [psobject]$StandardPropList,
        [Parameter(Position=10, Mandatory=$false)] [psobject]$ExPropList,
        [Parameter(Position=11, Mandatory=$false)] [switch]$ShowRequest,
        [Parameter(Position=12, Mandatory=$false)] [String]$SaveToSentItems,
        [Parameter(Position=13, Mandatory=$false)] [switch]$SendMail,
        [Parameter(Position=14, Mandatory=$false)] [psobject]$ReplyTo,
        [Parameter(Position=17, Mandatory=$false)] [bool]$RequestReadRecipient,
        [Parameter(Position=18, Mandatory=$false)] [bool]$RequestDeliveryRecipient
    )
    Begin{
        $NewMessage = "{" + "`r`n"
        if($SendMail.IsPresent){
             $NewMessage += "  `"Message`" : {" + "`r`n"
        }
        if(![String]::IsNullOrEmpty($Subject)){
            $NewMessage +=  "`"Subject`": `"" + $Subject + "`"" + "`r`n"
        }
        if($SenderEmailAddress -ne $null){
             if($NewMessage.Length -gt 5){$NewMessage += ","}
            $NewMessage +=  "`"Sender`":{" + "`r`n"
            $NewMessage +=  " `"EmailAddress`":{" + "`r`n"
            $NewMessage +=  "  `"Name`":`"" + $SenderEmailAddress.Name + "`"," + "`r`n"
            $NewMessage +=  "  `"Address`":`"" + $SenderEmailAddress.Address + "`"" + "`r`n"
            $NewMessage +=  "}}" + "`r`n"
        }
        if(![String]::IsNullOrEmpty($Body)){
             if($NewMessage.Length -gt 5){$NewMessage += ","}
            $NewMessage +=  "`"Body`": {"+ "`r`n"
            $NewMessage +=  "`"ContentType`": `"HTML`"," + "`r`n"
            $NewMessage +=  "`"Content`": `"" + $Body + "`"" + "`r`n"
            $NewMessage +=  "}" + "`r`n"
        }      
        
        $toRcpcnt = 0;
        if($ToRecipients -ne $null){
             if($NewMessage.Length -gt 5){$NewMessage += ","}
            $NewMessage +=  "`"ToRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $ToRecipients) {
                if($toRcpcnt -gt 0){
                    $NewMessage +=  "      ,{ "+ "`r`n"   
                }
                else{
                    $NewMessage +=  "      { "+ "`r`n"
                }           
                $NewMessage +=  " `"EmailAddress`":{" + "`r`n"
                $NewMessage +=  "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage +=  "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage +=  "}}" + "`r`n"
                $toRcpcnt++
            }
            $NewMessage +=  "  ]" + "`r`n"  
        }
        $ccRcpcnt = 0
        if($CcRecipients -ne $null){
             if($NewMessage.Length -gt 5){$NewMessage += ","}
            $NewMessage +=  "`"CcRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $CcRecipients) {
                if($ccRcpcnt  -gt 0){
                    $NewMessage +=  "      ,{ "+ "`r`n"   
                }
                else{
                    $NewMessage +=  "      { "+ "`r`n"
                }           
                $NewMessage +=  " `"EmailAddress`":{" + "`r`n"
                $NewMessage +=  "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage +=  "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage +=  "}}" + "`r`n"
                $ccRcpcnt++
            }
            $NewMessage +=  "  ]" + "`r`n"  
        }
        $bccRcpcnt = 0
        if($bccRecipients -ne $null){
             if($NewMessage.Length -gt 5){$NewMessage += ","}
            $NewMessage +=  "`"BccRecipients`": [ " + "`r`n"
            foreach ($EmailAddress in $bccRecipients) {
                if($bccRcpcnt -gt 0){
                    $NewMessage +=  "      ,{ "+ "`r`n"   
                }
                else{
                    $NewMessage +=  "      { "+ "`r`n"
                }           
                $NewMessage +=  " `"EmailAddress`":{" + "`r`n"
                $NewMessage +=  "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage +=  "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage +=  "}}" + "`r`n"
                $bccRcpcnt++
            }
            $NewMessage +=  "  ]" + "`r`n"  
        }
        $ReplyTocnt = 0
        if($ReplyTo -ne $null){
            if($NewMessage.Length -gt 5){$NewMessage += ","}
            $NewMessage +=  "`"ReplyTo`": [ " + "`r`n"
            foreach ($EmailAddress in $ReplyTo) {
                if($ReplyTocnt -gt 0){
                    $NewMessage +=  "      ,{ "+ "`r`n"   
                }
                else{
                    $NewMessage +=  "      { "+ "`r`n"
                }           
                $NewMessage +=  " `"EmailAddress`":{" + "`r`n"
                $NewMessage +=  "  `"Name`":`"" + $EmailAddress.Name + "`"," + "`r`n"
                $NewMessage +=  "  `"Address`":`"" + $EmailAddress.Address + "`"" + "`r`n"
                $NewMessage +=  "}}" + "`r`n"
                $ReplyTocnt++
            }
            $NewMessage +=  "  ]" + "`r`n"  
        }
        if($RequestDeliveryRecipient){
            $NewMessage +=  ",`"IsDeliveryReceiptRequested`": true`r`n"
        }
        if($RequestReadRecipient){
            $NewMessage +=  ",`"IsReadReceiptRequested`": true `r`n"
        }
        if($StandardPropList -ne $null){
            foreach ($StandardProp in $StandardPropList) {
                if($NewMessage.Length -gt 5){$NewMessage += ","}
                switch($StandardProp.PropertyType){
                    "Single" {
                        if($StandardProp.QuoteValue){
                           $NewMessage +=  "`"" + $StandardProp.Name + "`": `"" + $StandardProp.Value + "`"" + "`r`n" 
                        }
                        else{
                            $NewMessage +=  "`"" + $StandardProp.Name + "`": " + $StandardProp.Value +  "`r`n" 
                        }
                        
                   
                   }
                    "Object"  {
                              if($StandardProp.isArray){
                                    $NewMessage +=  "`"" + $StandardProp.PropertyName + "`": [ {"+ "`r`n"
                              }else{
                                    $NewMessage +=  "`"" + $StandardProp.PropertyName + "`": {"+ "`r`n"
                              }                              
                              $acCount = 0
                              foreach ($PropKeyValue in $StandardProp.PropertyList) {
                                    if($acCount -gt 0){
                                        $NewMessage += ","
                                    }
                                    $NewMessage +=  "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Name + "`"" + "`r`n"
                                    $acCount++
                              }
                               if($StandardProp.isArray){
                                    $NewMessage +=  "}]" + "`r`n"
                               }else{
                                    $NewMessage +=  "}" + "`r`n"
                               }
                              
                    }
                    "ObjectCollection" {
                              if($StandardProp.isArray){
                                    $NewMessage +=  "`"" + $StandardProp.PropertyName + "`": ["+ "`r`n"
                              }else{
                                    $NewMessage +=  "`"" + $StandardProp.PropertyName + "`": {"+ "`r`n"
                              }     
                              foreach ($EnclosedStandardProp in $StandardProp.PropertyList) {
                                  $NewMessage +=  "`"" + $EnclosedStandardProp.PropertyName + "`": {"+ "`r`n"
                                   foreach ($PropKeyValue in $EnclosedStandardProp.PropertyList) {
                                       $NewMessage +=  "`"" + $PropKeyValue.Name + "`": `"" + $PropKeyValue.Name + "`"," + "`r`n"
                                  }
                                  $NewMessage +=  "}" + "`r`n"
                              }
                               if($StandardProp.isArray){
                                    $NewMessage +=  "]" + "`r`n"
                               }else{
                                    $NewMessage +=  "}" + "`r`n"
                               }
                    }
                   
                }
            }
        }                  
        $atcnt = 0
        if($Attachments -ne $null){
             if($NewMessage.Length -gt 5){$NewMessage += ","}
            $NewMessage +=  "  `"Attachments`": [ " + "`r`n"
            foreach ($Attachment in $Attachments) {
                $Item = Get-Item $Attachment
                if($atcnt -gt 0)
                {
                     $NewMessage +=  "   ,{" + "`r`n"                       
                }
                else{
                     $NewMessage +=  "    {" + "`r`n"
                }               
                $NewMessage +=  "     `"@odata.type`": `"#Microsoft.OutlookServices.FileAttachment`"," + "`r`n"
                $NewMessage +=  "     `"Name`": `"" + $Item.Name + "`"," + "`r`n"
                $NewMessage +=  "     `"ContentBytes`": `" " + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Attachment)) + "`"" + "`r`n"
                $NewMessage +=  "    } " + "`r`n"
                $atcnt++
            }
            $NewMessage +=  "  ]" + "`r`n"
        }        
        if($ExPropList -ne $null){
             if($NewMessage.Length -gt 5){$NewMessage += ","}
            $NewMessage +=  "`"SingleValueExtendedProperties`": ["+ "`r`n"
            $propCount = 0
            foreach($Property in $ExPropList){
               if($propCount -eq 0){
                   $NewMessage +=  "{"+ "`r`n"
               }
               else{
                   $NewMessage +=  ",{"+ "`r`n"
               }
               if($Property.PropertyType -eq "Tagged"){
                     $NewMessage +=  "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Id + "`", " + "`r`n"
               }
               else{
                   if($Property.Type -eq "String"){
                       $NewMessage +=  "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Guid + " Name " + $Property.Id + "`", " + "`r`n"
                   }
                   else{
                       $NewMessage +=  "`"PropertyId`":`"" + $Property.DataType + " " + $Property.Guid + " Id " + $Property.Id + "`", " + "`r`n"
                   }
               }
               $NewMessage +=  "`"Value`":`"" + $Property.Value + "`""+ "`r`n"
               $NewMessage +=  " } " + "`r`n"
               $propCount++
            }
            $NewMessage +=  "]" + "`r`n"   
        } 
        if(![String]::IsNullOrEmpty($SaveToSentItems)){
            $NewMessage += "}   ,`"SaveToSentItems`": `"" + $SaveToSentItems.ToLower() + "`""+ "`r`n"
        }   
        $NewMessage +=  "}"
        if($ShowRequest.IsPresent){
            Write-Host $NewMessage
        }
        return, $NewMessage
    }
}
#endregion
function HexStringToByteArray($HexString)
{
	$ByteArray =  New-Object Byte[] ($HexString.Length/2);
  	for ($i = 0; $i -lt $HexString.Length; $i += 2)
	{
		 $ByteArray[$i/2] = [Convert]::ToByte($HexString.Substring($i, 2), 16)
	} 
 	Return @(,$ByteArray)

}