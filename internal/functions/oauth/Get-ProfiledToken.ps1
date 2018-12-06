function Get-ProfiledToken {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $ResourceURL
    )
    Process {
        if ([String]::IsNullOrEmpty($ResourceURL)) {
            $ResourceURL = "graph.microsoft.com" 
        }
        if ($Script:TokenCache.ContainsKey($ResourceURL)) {
            $AccessToken = $null
            if ([String]::IsNullOrEmpty($MailboxName)) {
                $firstToken = $Script:TokenCache[$ResourceURL].GetEnumerator() | select -first 1
                $AccessToken = $firstToken.Value
            }
            else {
                $HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
                if ($Script:TokenCache[$ResourceURL].ContainsKey($HostDomain)) {				
                    $AccessToken = $Script:TokenCache[$ResourceURL][$HostDomain]
                }
            }
            if ($AccessToken -ne $null) {
                $MailboxName = $AccessToken.mailbox
                if ($AccessToken.ADAL) {
					$Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($AccessToken.Authority)
					$token = ($Context.AcquireTokenSilentAsync($AccessToken.resource,$AccessToken.clientid)).Result
					if($token.AccessToken -ne $AccessToken.AccessToken){
						write-host "Refreshed Token ADAL"
						if ([bool]($token.PSobject.Properties.name -match "AccessToken")) {
							#$AdalToken.access_token = 
							Add-Member -InputObject $token -NotePropertyName access_token -NotePropertyValue (Get-ProtectedToken -PlainToken $token.AccessToken) -Force
						}
						Add-Member -InputObject $token -NotePropertyName clientid -NotePropertyValue $AccessToken.clientid
						Add-Member -InputObject $token -NotePropertyName ADAL -NotePropertyValue $True
						Add-Member -InputObject $token -NotePropertyName redirectUrl -NotePropertyValue $AccessToken.redirectUrl
						Add-Member -InputObject $token -NotePropertyName resource -NotePropertyValue $AccessToken.resource
						Add-Member -InputObject $token -NotePropertyName mailbox -NotePropertyValue $AccessToken.mailbox
						Add-Member -InputObject $token -NotePropertyName resourceCache -NotePropertyValue $AccessToken.resourceCache
						if(![String]::IsNullOrEmpty($AccessToken.TenantId)){
							Add-Member -InputObject $token -NotePropertyName TenantId -NotePropertyValue $AccessToken.TenantId -Force
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
                else {
                    #Check for expired Token
                    $minTime = new-object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc);
                    $expiry = $minTime.AddSeconds($AccessToken.expires_on)
                    if ($expiry -le [DateTime]::Now.ToUniversalTime().AddMinutes(10)) {
                        if ([bool]($AccessToken.PSobject.Properties.name -match "refresh_token")) {
                            write-host "Refresh Token"
                            $AccessToken = Invoke-RefreshAccessToken -MailboxName $MailboxName -AccessToken $AccessToken
                        }
                        else {
                            throw "App Token has expired"
                        }
					
                    }
                }
                return $AccessToken
            }
			
        }

    }
}