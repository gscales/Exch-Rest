function Get-EXRAccessTokenADAL {
    [CmdletBinding()]
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
        $Prompt,

        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $CacheCredentials,

        [Parameter(Position = 8, Mandatory = $false)]
        [string]
        $TenantId,

        [Parameter(Position = 9, Mandatory = $false)]
        [switch]
        $useLoggedOnCredentials,
		
        [Parameter(Position = 10, Mandatory = $false)]
        [String]
        $AADUserName
		
    )
    Begin {
		Add-Type -AssemblyName System.Web
        if ([String]::IsNullOrEmpty($Prompt)) {
            $Prompt = "RefreshSession"
        }
        $PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList $Prompt
        if ([String]::IsNullOrEmpty($redirectUrl)) {
            $redirectUrl = [System.Web.HttpUtility]::UrlEncode("urn:ietf:wg:oauth:2.0:oob")
        }
        $ResourceURI = "https://" + $ResourceURL
        $DomainName = $MailboxName.Split('@')[1]
        $EndpointUri = 'https://login.microsoftonline.com/' + (Get-EXRTenantId -DomainName $DomainName)
        $Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($EndpointUri)
        $Script:ADALContext = $EndpointUri
        if ($useLoggedOnCredentials.IsPresent) {
            $AADCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential" -ArgumentList $AADUserName
            $authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($Context, $ResourceURI, $ClientId, $AADcredential)
            if ($authResult.Result.AccessToken) {
                $token = $authResult.Result
            }
            elseif ($authResult.Exception) {    
                throw "An error occured getting access token: $($authResult.Exception.InnerException)"    
            }
        }
        else {
            if (![String]::IsNullOrEmpty(($ClientSecret))) {
                $ClientCredentails = new-object Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential -ArgumentList $ClientId,$ClientSecret
                $authResult = $Context.AcquireTokenAsync($ResourceURI, $ClientCredentails)
                
            }else{
			    $authResult = $Context.AcquireTokenAsync($ResourceURI, $ClientId, $redirectUrl, $PromptBehavior)
            }

			if ($authResult.Result.AccessToken) {
                $token = $authResult.Result
            }
            elseif ($authResult.Exception) {    
                throw "An error occured getting access token: $($authResult.Exception.InnerException)"    
            }
        }
        if ($token) {
            if ([bool]($token.PSobject.Properties.name -match "AccessToken")) {
                #$AdalToken.access_token = 
                Add-Member -InputObject $Token -NotePropertyName access_token -NotePropertyValue (Get-ProtectedToken -PlainToken $token.AccessToken) -Force
            }
            Add-Member -InputObject $token -NotePropertyName clientid -NotePropertyValue $ClientId
            Add-Member -InputObject $token -NotePropertyName ADAL -NotePropertyValue $True
            Add-Member -InputObject $token -NotePropertyName redirectUrl -NotePropertyValue $redirectUrl
            Add-Member -InputObject $token -NotePropertyName resource -NotePropertyValue $ResourceURI
            Add-Member -InputObject $token -NotePropertyName resourceCache -NotePropertyValue $ResourceURL
            Add-Member -InputObject $token -NotePropertyName mailbox -NotePropertyValue $MailboxName
            if (![String]::IsNullOrEmpty(($ClientSecret))) {
                Add-Member -InputObject $token -NotePropertyName refresh -NotePropertyValue $false
            }else{
                 Add-Member -InputObject $token -NotePropertyName refresh -NotePropertyValue $true
            }
            if (![String]::IsNullOrEmpty($TenantId)) {
                Add-Member -InputObject $token -NotePropertyName TenantId -NotePropertyValue $TenantId
            }
            if ($Beta.IsPresent) {
                Add-Member -InputObject $token -NotePropertyName Beta -NotePropertyValue $True
            }
            if ($CacheCredentials.IsPresent) {
                if (!$Script:TokenCache.ContainsKey($ResourceURL)) {	
                    $ResourceTokens = @{}		
                    $Script:TokenCache.Add($ResourceURL, $ResourceTokens)
                }
                Add-Member -InputObject $token -NotePropertyName Cached -NotePropertyValue $true				
                $HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
                if (!$Script:TokenCache[$ResourceURL].ContainsKey($HostDomain)) {			
                    $Script:TokenCache[$ResourceURL].Add($HostDomain, $token)
                }
                else {
                    $Script:TokenCache[$ResourceURL][$HostDomain] = $token
                }
                write-host ("Cached Token for " + $ResourceURL + " " + $HostDomain)
            }
        }
        return $token
    }
}
